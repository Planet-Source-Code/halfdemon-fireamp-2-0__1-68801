VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMediaLib 
   Caption         =   "FireAMP Media Library"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMediaLib.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7575
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10815
      Begin VB.PictureBox picBarBack 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3360
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   287
         TabIndex        =   5
         Top             =   3480
         Width           =   4335
         Begin VB.PictureBox picBar 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   0
            Picture         =   "frmMediaLib.frx":08CA
            ScaleHeight     =   24
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   37
            TabIndex        =   6
            Top             =   0
            Width           =   555
         End
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3360
         Picture         =   "frmMediaLib.frx":138C
         Top             =   2880
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Media library loading ..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4320
         TabIndex        =   4
         Top             =   3000
         Width           =   2925
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      Left            =   4560
      TabIndex        =   7
      Top             =   360
      Width           =   4455
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5280
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediaLib.frx":1C56
            Key             =   "List"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediaLib.frx":1FA8
            Key             =   "Music"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediaLib.frx":22FA
            Key             =   "Block"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstMediaLib 
      Height          =   5535
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9763
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "SngTitle"
         Text            =   "Songtitle"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Artist"
         Text            =   "Artist"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Album"
         Text            =   "Album"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Year"
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Comment"
         Text            =   "Comment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Track"
         Text            =   "Track #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Filename"
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TreeView trvMediaLib 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   10821
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Media Library"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   2
      Top             =   7080
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      FillColor       =   &H80000010&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   -1440
      Top             =   6360
      Width           =   12135
   End
End
Attribute VB_Name = "frmMediaLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
' defunct

Private Sub Command1_Click()
Unload Me
End Sub


Private Sub Combo1_Click()
 combo1_KeyDown vbKeyReturn, 0
End Sub

Private Sub combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Combo1.AddItem Combo1.Text
 
If KeyCode = vbKeyReturn Then
Dim i As Integer, LST As ListItem

Dim t As tagID3_1x
lstMediaLib.ListItems.Clear

For i = 0 To frmFirePL.lstPaths.ListCount - 1

If getFileExtensionFromPath(frmFirePL.lstPaths.List(i)) = "mp3" Then
t = readID3_1x(frmFirePL.lstPaths.List(i))

Combo1.Text = Trim(Combo1.Text)

If InStr(t.Album, Combo1.Text) Or InStr(t.Artist, Combo1.Text) Or InStr(t.Title, Combo1.Text) Or InStr(getGenre(t.Genre), Combo1.Text) Then
Set LST = lstMediaLib.ListItems.Add(, , t.Title)
LST.SubItems(1) = t.Artist
LST.SubItems(2) = t.Album
LST.SubItems(3) = t.Year
LST.SubItems(4) = t.Comment
LST.SubItems(5) = t.Track
LST.SubItems(6) = frmFirePL.lstPaths.List(i)
LST.Key = LST.SubItems(6)
LST.ToolTipText = toStdString(t.Artist) & " / " & toStdString(t.Title) & " / " & toStdString(t.Album)
End If

End If
Next i
lblInfo.Caption = "Search: " & lstMediaLib.ListItems.Count & " Clip(s) found."

End If

End If
End Sub

Private Sub Form_Load()

lib.RemoveAll
invLib.RemoveAll
Show
'On Error Resume Next
Dim trvMediaNode As Node
Set trvMediaNode = trvMediaLib.Nodes.Add(, tvwNext, "Lib", "Library", imgList.ListImages(3).Key)
Set trvMediaNode = trvMediaLib.Nodes.Add(, tvwNext, "Play", "Playlist(s)", imgList.ListImages(3).Key)
Set trvMediaNode = trvMediaLib.Nodes.Add("Play", tvwChild, "usrPlay1", "Playlist1", imgList.ListImages(1).Key)
Set trvMediaNode = trvMediaLib.Nodes.Add("Play", tvwChild, "usrPlay2", "Playlist2", imgList.ListImages(1).Key)
Set trvMediaNode = trvMediaLib.Nodes.Add("Play", tvwChild, "usrPlay3", "Playlist3", imgList.ListImages(1).Key)

Set trvMediaNode = trvMediaLib.Nodes.Add(, tvwNext, "Album", "Album(s)", imgList.ListImages(3).Key)
Set trvMediaNode = trvMediaLib.Nodes.Add(, tvwNext, "Artist", "Artist(s)", imgList.ListImages(3).Key)

Dim i As Integer, LST As ListItem, j As Integer, k As Integer, l As Integer
Dim file As String

Dim alb As New Dictionary, art As New Dictionary
DoEvents

For i = 0 To frmFirePL.lstPaths.ListCount - 1
file = frmFirePL.lstPaths.List(i)
Dim t As tagID3_1x

If getFileExtensionFromPath(file) = "mp3" Then
t = readID3_1x(file)

Set LST = lstMediaLib.ListItems.Add(, file, toStdString(t.Title))
LST.SubItems(1) = toStdString(t.Artist)
LST.SubItems(2) = toStdString(t.Album)
LST.SubItems(3) = toStdString(t.Year)
LST.SubItems(4) = toStdString(t.Comment)
LST.SubItems(5) = t.Track
LST.SubItems(6) = file
LST.ToolTipText = toStdString(t.Artist) & " / " & toStdString(t.Title) & " / " & toStdString(t.Album)
If Not alb.Exists(Trim(t.Album)) Then
Set trvMediaNode = trvMediaLib.Nodes.Add("Album", tvwChild, "album" & CStr(k), Trim(t.Album), imgList.ListImages(2).Key)
alb.Add Trim(t.Album), Trim(t.Album)
k = k + 1
End If

If Not art.Exists(Trim(t.Artist)) Then
Set trvMediaNode = trvMediaLib.Nodes.Add("Artist", tvwChild, "artist" & CStr(l), Trim(t.Artist))
art.Add Trim(t.Artist), Trim(t.Artist)
l = l + 1
End If
 
Else
Set LST = lstMediaLib.ListItems.Add(, file, getFileTitleFromPath(file))
LST.SubItems(2) = getFileExtensionFromPath(file) & " clip"
LST.SubItems(6) = file
 
End If
lib.Add file, i + 1
invLib.Add i + 1, file
DoEvents
updateBar picBar, picBarBack, frmFirePL.lstPaths.ListCount - 1, i

Next i
Frame1.Visible = False
End Sub


Private Sub Form_Resize()
If WindowState <> vbMinimized Then
trvMediaLib.Width = Width / 3 - 10
lstMediaLib.Width = 2 * Width / 3 - 100
trvMediaLib.Left = 20
lstMediaLib.Left = trvMediaLib.Width
trvMediaLib.Top = 0
lstMediaLib.Top = Combo1.Height
Combo1.Top = 0
Combo1.Left = Width - trvMediaLib.Width - Combo1.Width / 2
Combo1.Left = Combo1.Left
Combo1.Top = 0
trvMediaLib.Height = Height * (8 / 10)
lstMediaLib.Height = Height * (8 / 10) - Combo1.Height + 30
Shape1.Top = trvMediaLib.Height - 20
Shape1.Height = Height / 5
Shape1.Width = 3 * Width
lblInfo.Top = Shape1.Top + Shape1.Height / 2 - lblInfo.Height / 2
End If
End Sub

Private Sub lstMediaLib_DblClick()
curFile = lstMediaLib.SelectedItem.Key
frmFirePL.lstPl.ListItems(lib.Item(curFile)).Selected = True
frmFirePL.lstPaths.ListIndex = frmFirePL.lstPl.SelectedItem.Index - 1
frmFireMain.picBtn_Click 1 'stop
frmFireMain.picBtn_Click 0 'play
End Sub


Private Sub trvMediaLib_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer, LST As ListItem

Dim t As tagID3_1x
lstMediaLib.ListItems.Clear

For i = 0 To frmFirePL.lstPaths.ListCount - 1

If getFileExtensionFromPath(frmFirePL.lstPaths.List(i)) = "mp3" Then
t = readID3_1x(frmFirePL.lstPaths.List(i))

If Trim(Node.Text) = Trim(t.Album) Or Trim(Node.Text) = Trim(t.Artist) Then
Set LST = lstMediaLib.ListItems.Add(, , t.Title)
LST.SubItems(1) = t.Artist
LST.SubItems(2) = t.Album
LST.SubItems(3) = t.Year
LST.SubItems(4) = t.Comment
LST.SubItems(5) = t.Track
LST.SubItems(6) = frmFirePL.lstPaths.List(i)
LST.Key = LST.SubItems(6)
LST.ToolTipText = toStdString(t.Artist) & " / " & toStdString(t.Title) & " / " & toStdString(t.Album)
End If

End If
Next i
lblInfo.Caption = lstMediaLib.ListItems.Count & " Clip(s) "
End Sub

