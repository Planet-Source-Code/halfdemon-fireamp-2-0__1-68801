VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmFirePL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Playlist"
   ClientHeight    =   6345
   ClientLeft      =   2070
   ClientTop       =   2055
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3120
      Left            =   3840
      Picture         =   "Playlist.frx":0000
      ScaleHeight     =   208
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   4
      Top             =   480
      Width           =   150
      Begin VB.PictureBox picBar 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   15
         Picture         =   "Playlist.frx":1A42
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   7
         TabIndex        =   5
         Top             =   120
         Width           =   105
      End
   End
   Begin VB.ListBox lstPaths 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   5640
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox picSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4965
      Left            =   0
      Picture         =   "Playlist.frx":1C04
      ScaleHeight     =   331
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   287
      TabIndex        =   0
      Top             =   0
      Width           =   4305
      Begin MSComctlLib.ListView lstPl 
         Height          =   3555
         Left            =   480
         TabIndex        =   1
         ToolTipText     =   "The Playlist"
         Top             =   960
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   6271
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         OLEDragMode     =   1
         OLEDropMode     =   1
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   12632256
         BackColor       =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Length"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "PlayList"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmFirePL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' the playlist
'

' last updated: 2006 May 05, Humanoid

' some variables for scroll bars
Private yy As Long, temp As Long


Private Sub Form_Load()
Width = picSkin.Width
Height = picSkin.Height
Dim windowRegion As Long
    windowRegion = MakeRegion(picSkin)
    SetWindowRgn Me.hwnd, windowRegion, True

lstPl.ColumnHeaders.Item(1).Width = CInt(lstPl.Width / 1.3)
lstPl.ColumnHeaders.Item(2).Width = lstPl.Width - lstPl.ColumnHeaders.Item(1).Width


Left = frmFireMain.Left + frmFireMain.Width



End Sub

Private Sub lstPL_Click()
On Error GoTo e:
    picBar.Top = (picBack.ScaleHeight - picBar.Height) * (lstPl.SelectedItem.Index / lstPl.ListItems.Count)
Exit Sub
e:
End Sub

Public Sub lstPl_DblClick()
On Error GoTo e
Dim i As Integer
    frmFireMain.picBtn_Click 1
    curFile = lstPaths.List(lstPl.SelectedItem.Index - 1)
    frmFirePL.lstPaths.ListIndex = lstPl.SelectedItem.Index - 1
    frmFireMain.picBtn_Click 0
    
    
playingIndex = lstPl.SelectedItem.Index
Exit Sub
e:
End Sub

Private Sub lstPL_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo e:
    picBar.Top = (picBack.ScaleHeight - picBar.Height) * (lstPl.SelectedItem.Index / lstPl.ListItems.Count)
    
    If KeyCode = 13 Then 'enter key
    frmFirePL.lstPaths.ListIndex = lstPl.SelectedItem.Index - 1
     curFile = lstPaths.List(lstPl.SelectedItem.Index - 1)
     frmFireMain.picBtn_Click 1 'stop
     frmFireMain.picBtn_Click 0 'play
    End If

Exit Sub
e:
End Sub

Private Sub lstPl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbRightButton Then

On Error Resume Next
If lstPaths.ListCount > 0 Then
 frmDummy.mnuViewInfo.Enabled = True

    If getFileExtensionFromPath(lstPaths.List(lstPl.SelectedItem.Index - 1)) <> "mp3" Then
            frmDummy.mnuTagEdit.Enabled = False
            frmDummy.mnuGoogleSearch.Enabled = False
    Else
            frmDummy.mnuTagEdit.Enabled = True
            frmDummy.mnuGoogleSearch.Enabled = True

            Dim tempTag As tagID3_1x, isFilled As Boolean
            tempTag = readID3_1x(lstPaths.List(lstPl.SelectedItem.Index - 1))

                If tempTag.Album = "" Then
                    frmDummy.mnuSearchForAlbum.Visible = False
                Else
                    frmDummy.mnuSearchForAlbum.Visible = True
                    frmDummy.mnuSearchForAlbum.Caption = Trim(tempTag.Album)
                End If

                If tempTag.Artist = "" Then
                    frmDummy.mnuSearchArtist.Visible = False
                Else
                    frmDummy.mnuSearchArtist.Visible = True
                    frmDummy.mnuSearchArtist.Caption = Trim(tempTag.Artist)
                End If
     End If

Else
 frmDummy.mnuViewInfo.Enabled = False
 frmDummy.mnuTagEdit.Enabled = False

End If
        Me.PopupMenu frmDummy.mnuPlaylist
End If

End Sub


Private Sub lstPl_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Dim i As Integer

For i = 1 To Data.Files.Count
If isMediaFile(Data.Files.Item(i)) Then
lstPaths.AddItem Data.Files.Item(i)
lstPl.ListItems.Add , , getFileTitleFromPath(Data.Files.Item(i))
End If
Next

End Sub

Private Sub picBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 1 Then
If picBar.Top + picBar.Height > picBack.ScaleHeight And yy < y Then Exit Sub
If picBar.Top < 10 And yy > y Then Exit Sub

If temp < lstPl.ListItems.Count / 2 Then
temp = ((picBar.Top) * lstPl.ListItems.Count) \ picBack.ScaleHeight

Else
temp = ((picBar.Top + picBar.Height) * lstPl.ListItems.Count) \ picBack.ScaleHeight

End If

lstPl.ListItems(temp + 1).EnsureVisible


If Abs(yy - y) > 50 Then Exit Sub
       picBar.Top = picBar.Top + y
       yy = y
       
       
End If


End Sub


Private Sub picBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If picBar.Top + picBar.Height > picBack.ScaleHeight And yy < y Then picBar.Top = picBack.ScaleHeight - picBar.Height
If picBar.Top < 10 And yy > y Then picBar.Top = 0

End Sub

Private Sub picSkin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
      ReleaseCapture
      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

