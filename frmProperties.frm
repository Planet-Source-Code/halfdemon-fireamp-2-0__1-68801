VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   6360
      Width           =   1335
   End
   Begin TabDlg.SSTab tabProperties 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "File   "
      TabPicture(0)   =   "frmProperties.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblVideoSize"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblBitRate"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblVideoCodec"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblAudioCodec"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblMediaType"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblSize"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblLength"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtPath"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Content"
      TabPicture(1)   =   "frmProperties.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblGenre"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblYear"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblAlbum"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblArtist"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblTitle"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label16"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label15"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label14"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label13"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label11"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label10"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label7"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtDescription"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdEditTags"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.CommandButton cmdEditTags 
         Caption         =   "&Edit Tags"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -71400
         TabIndex        =   4
         Top             =   5520
         Width           =   1455
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   -74640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   4320
         Width           =   4695
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   1215
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   4680
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
         Height          =   195
         Left            =   360
         TabIndex        =   32
         Top             =   4680
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Length:"
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Size:"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblLength 
         Caption         =   "-"
         Height          =   255
         Left            =   1680
         TabIndex        =   29
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblSize 
         Caption         =   "-"
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "View Advanced details about the selected clip"
         Height          =   195
         Left            =   1065
         TabIndex        =   27
         Top             =   840
         Width           =   3285
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "View Content Information  of the selected clip"
         Height          =   195
         Left            =   -73920
         TabIndex        =   26
         Top             =   840
         Width           =   3225
      End
      Begin VB.Label Label8 
         Caption         =   "Media Type:"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblMediaType 
         Caption         =   "-"
         Height          =   255
         Left            =   1680
         TabIndex        =   24
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Audio Codec:"
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label lblAudioCodec 
         Caption         =   "-"
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "Video Codec:"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lblVideoCodec 
         Caption         =   "-"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Label Label9 
         Caption         =   "Bit rate:"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lblBitRate 
         Caption         =   "-"
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Video Size:"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label lblVideoSize 
         Caption         =   "-"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Title:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   15
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Artist:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   14
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Album:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   13
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Year:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   12
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Description:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   11
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Genre:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   10
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Caption         =   "-"
         Height          =   255
         Left            =   -73800
         TabIndex        =   9
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label lblArtist 
         Caption         =   "-"
         Height          =   255
         Left            =   -73800
         TabIndex        =   8
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label lblAlbum 
         Caption         =   "-"
         Height          =   255
         Left            =   -73800
         TabIndex        =   7
         Top             =   2400
         Width           =   3735
      End
      Begin VB.Label lblYear 
         Caption         =   "-"
         Height          =   255
         Left            =   -73800
         TabIndex        =   6
         Top             =   2760
         Width           =   3855
      End
      Begin VB.Label lblGenre 
         Caption         =   "-"
         Height          =   255
         Left            =   -73800
         TabIndex        =   5
         Top             =   3120
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'
' last updated: 2006 June 29
'

' shows file properties of the selected clip

' sub to read properties
Public Sub readProperties(FileName As String)
On Error GoTo CannotShowThisForm

Dim fGraph As New FilgraphManager ' filter graph manager
Dim filterInfo As IAMCollection   ' collection of filters used in the current file
Dim fStats As IMediaPosition      ' for length
Dim vStats As IBasicVideo2        ' for video information
Dim ppUnk As Object               ' temp


' render the file
fGraph.RenderFile FileName

Set filterInfo = fGraph.FilterCollection
Set fStats = fGraph


lblLength.Caption = "[" & modCommon.convertToStdTime(fStats.Duration) & "]"
txtPath.Text = FileName
lblSize.Caption = Round(FileLen(FileName) / (1024& * 1024&), 2) & " MB"
lblTitle.Caption = getFileTitleFromPath(FileName)

Select Case getFileExtensionFromPath(FileName)

' Audio
Case "mp3", "mp2", "mp1", "wav", "wma", "mid", "rmi"

  lblMediaType.Caption = "Audio"
  
  ' read filter info and store in ppUnk
  ' generally for audio ,"item" :
  ' 0: audio device
  ' 1: audio codec
  ' 2: audio renderer
  ' 3: file path
  
  filterInfo.Item 1, ppUnk
  
  lblAudioCodec.Caption = ppUnk.Name
  lblVideoCodec.Caption = "-"
  lblBitRate.Caption = "-"
  lblVideoSize.Caption = "-"
  
If getFileExtensionFromPath(FileName) = "mp3" Then
 Dim tag As tagID3_1x
 tag = readID3_1x(FileName)
 lblTitle.Caption = tag.Title
 lblAlbum.Caption = tag.Album
 lblArtist.Caption = tag.Artist
 txtDescription.Text = tag.Comment
 lblGenre.Caption = getGenre(tag.Genre)
 lblYear.Caption = tag.Year
 lblBitRate.Caption = ReadMP3Header(FileName).BitRate & "Kbps"
 cmdEditTags.Enabled = True
End If

Case "mpeg", "mpg", "mov", "mpe", "rm", "wmv"

Set vStats = fGraph

  lblMediaType.Caption = "Video"
  lblBitRate.Caption = "-"
  
  ' video size
  lblVideoSize.Caption = vStats.VideoWidth & " x " & vStats.VideoHeight
Set vStats = Nothing
  filterInfo.Item 2, ppUnk
  
  ' generally for video ,"item" :
  ' 0: audio device
  ' 1: video device
  ' 2: audio codec
  ' 3: video codec
  ' 4: video splitter
  ' 5: file path
  
  ' audio, video codecs are interchanged sometimes
  If InStr(1, ppUnk.Name, "Audio") Then
  lblAudioCodec.Caption = ppUnk.Name
  Else
  lblVideoCodec.Caption = ppUnk.Name
  End If
  
  filterInfo.Item 3, ppUnk
  If InStr(1, ppUnk.Name, "Video") Then
  lblVideoCodec.Caption = ppUnk.Name
  Else
  lblAudioCodec.Caption = ppUnk.Name
  End If
  
Case "avi"
Set vStats = fGraph

  lblMediaType.Caption = "Video"
  lblBitRate.Caption = "-"
  lblVideoSize.Caption = vStats.VideoWidth & " x " & vStats.VideoHeight
Set vStats = Nothing

    
  filterInfo.Item 2, ppUnk
  If InStr(1, ppUnk.Name, "video") Then
  lblVideoCodec.Caption = ppUnk.Name
  Else
  lblVideoCodec.Caption = "-"
  End If
  
   filterInfo.Item 3, ppUnk
  If InStr(1, ppUnk.Name, "Audio") Then
  lblAudioCodec.Caption = ppUnk.Name
  Else
  lblAudioCodec.Caption = "-"
  End If
  
  
Case Else
  lblMediaType.Caption = "-"
End Select



Set fGraph = Nothing
Set filterInfo = Nothing
Set fStats = Nothing
Set ppUnk = Nothing

Exit Sub
CannotShowThisForm:
 Beep
Hide
 
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEditTags_Click()

frmTagEditor.importCurrentPlaylist
frmTagEditor.lstFiles.ListIndex = frmFirePL.lstPl.SelectedItem.Index - 1
frmTagEditor.Show
End Sub

Private Sub txtPath_Click()
'ShellExecute hwnd, "", "", txtPath.Text, "", 0
End Sub
