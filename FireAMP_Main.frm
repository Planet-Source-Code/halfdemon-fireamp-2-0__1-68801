VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmFireMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "FireAMP!"
   ClientHeight    =   9120
   ClientLeft      =   915
   ClientTop       =   690
   ClientWidth     =   13890
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00584D37&
   Icon            =   "FireAMP_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrVisUpdate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   7200
   End
   Begin VB.PictureBox picCtrlSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   4320
      Picture         =   "FireAMP_Main.frx":08CA
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   12
      Top             =   6480
      Visible         =   0   'False
      Width           =   450
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7920
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open Media"
      Filter          =   $"FireAMP_Main.frx":13D4
   End
   Begin VB.PictureBox picBtnSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   6600
      Picture         =   "FireAMP_Main.frx":14CC
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   5040
      Width           =   900
   End
   Begin VB.PictureBox ScopeBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0E0FF&
      Height          =   1170
      Left            =   6360
      ScaleHeight     =   78
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   134
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   7200
   End
   Begin VB.Timer tmrPbr 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1680
      Top             =   7200
   End
   Begin VB.Timer tmrInfo 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   7200
   End
   Begin VB.PictureBox picSkin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5985
      Left            =   0
      Picture         =   "FireAMP_Main.frx":6970
      ScaleHeight     =   399
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   399
      TabIndex        =   0
      Top             =   0
      Width           =   5985
      Begin VB.PictureBox picPlaylist 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   4800
         ScaleHeight     =   3735
         ScaleWidth      =   5175
         TabIndex        =   23
         Top             =   5640
         Visible         =   0   'False
         Width           =   5175
         Begin VB.PictureBox picHide 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   5175
            TabIndex        =   26
            Top             =   3480
            Width           =   5175
         End
         Begin MSComctlLib.ListView lstPl 
            Height          =   3255
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "The Playlist"
            Top             =   480
            Width           =   4980
            _ExtentX        =   8784
            _ExtentY        =   5741
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            OLEDragMode     =   1
            OLEDropMode     =   1
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   12632256
            BackColor       =   0
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-Playlist-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   210
            Left            =   2040
            TabIndex        =   25
            Top             =   120
            Width           =   915
         End
      End
      Begin VB.PictureBox picScrollBuffer 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1000
         Left            =   840
         ScaleHeight     =   67
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   283
         TabIndex        =   22
         Top             =   2400
         Visible         =   0   'False
         Width           =   4245
      End
      Begin VB.PictureBox picTitle 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   337
         TabIndex        =   17
         Top             =   720
         Width           =   5055
         Begin VB.Label lblTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title Comes Here"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   435
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Visible         =   0   'False
            Width           =   3405
         End
      End
      Begin VB.PictureBox picAlbum 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   480
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   297
         TabIndex        =   16
         Top             =   1200
         Width           =   4455
         Begin VB.Label lblAlbum 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Burning Media Inc,"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   240
            Left            =   0
            TabIndex        =   18
            Top             =   0
            Width           =   1785
         End
      End
      Begin VB.PictureBox picCtrl 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   1
         Left            =   5040
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   4
         Top             =   180
         Width           =   180
      End
      Begin VB.PictureBox picCtrl 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   4800
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   3
         Top             =   180
         Width           =   180
      End
      Begin VB.PictureBox picScroll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   360
         ScaleHeight     =   345
         ScaleWidth      =   3645
         TabIndex        =   14
         Top             =   4440
         Width           =   3645
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FireAMP!"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   210
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   810
         End
      End
      Begin VB.PictureBox picBtn 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   2
         Left            =   4080
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   11
         ToolTipText     =   "Open Media ..."
         Top             =   5040
         Width           =   495
      End
      Begin VB.PictureBox picBtn 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   1
         Left            =   840
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   10
         ToolTipText     =   "Stop"
         Top             =   5040
         Width           =   495
      End
      Begin VB.PictureBox picBtn 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000C0C0&
         BorderStyle     =   0  'None
         Height          =   450
         Index           =   0
         Left            =   360
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   7
         ToolTipText     =   "Play / Pause"
         Top             =   5040
         Width           =   450
      End
      Begin VB.PictureBox picBarBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1440
         Picture         =   "FireAMP_Main.frx":7B802
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   168
         TabIndex        =   5
         ToolTipText     =   "Click to Seek"
         Top             =   5160
         Width           =   2520
         Begin VB.PictureBox picBarFront 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   120
            Picture         =   "FireAMP_Main.frx":7DDAC
            ScaleHeight     =   16
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   20
            TabIndex        =   6
            ToolTipText     =   "Drag to Seek"
            Top             =   15
            Width           =   300
         End
      End
      Begin VB.PictureBox fraDisplay 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   360
         ScaleHeight     =   3135
         ScaleWidth      =   5175
         TabIndex        =   19
         Top             =   1200
         Width           =   5175
         Begin VB.PictureBox picVisual 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            FillColor       =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            FontTransparent =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   2895
            Left            =   0
            ScaleHeight     =   193
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   345
            TabIndex        =   20
            ToolTipText     =   "Click to Configure this Visualization"
            Top             =   240
            Width           =   5175
            Begin VB.Label lblVis 
               Alignment       =   2  'Center
               BackColor       =   &H005C5CCD&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H00FFFFFF&
               Height          =   345
               Left            =   0
               TabIndex        =   21
               Top             =   0
               Visible         =   0   'False
               Width           =   4035
            End
         End
      End
      Begin VB.Frame fraVideo 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00 [00:00]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   4320
         TabIndex        =   2
         ToolTipText     =   "Current Time"
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FireAMP!"
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
         TabIndex        =   1
         Top             =   240
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmFireMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'
'================================================================================
' FireAMP! -- main
'================================================================================
'
'

' last updated: 2006 May 05, Humanoid


' variables to indicate if player is muted or paused
' or playing a new file
Private isMute As Boolean, isPlaying As Boolean, isNewFile As Boolean

' for startup time
Private t As Long

' temporary file to copy media except videos into
' for safer access
Private tFile As String

Private mx As Single
Private isMinus As Boolean, drawFlag As Boolean
Private p As Single, rFlag As Boolean
Private X As Integer, delay As Integer
' some conditional compilation variables
' use true for debug mode: no playlist, fast startup, default file play
' great for testing vis.True
#Const DBUG = False
' skin debug mode, does not load previous skin
#Const SkinDBUG = False

' sub: Form_Load
' form initialization procedures

Private Sub Form_Load()
DoEvents

Dim windowRegion As Long, i As Integer

t = Timer ' store startup time
'plColor = RGB(0, 64, 128)  ' default playlist forecolor
'plCurrent = RGB(0, 128, 255)
plColor = RGB(200, 200, 200)
plCurrent = vbWhite
'infoStr = "   Greetings, User!   *  Welcome to FireAMP!   "
 
Pos = 1
frmSplash.lblStatus.Caption = "Rendering Skin ..."
' skinning
Width = picSkin.Width
Height = picSkin.Height

    windowRegion = MakeRegion(picSkin)
    SetWindowRgn Me.hwnd, windowRegion, True
    
'setup buttons
isMute = False

' width and height of player buttons
Wx = picBtnSrc.ScaleWidth / 2
Wy = picBtnSrc.ScaleHeight / 4

' width and height of command buttons
cWx = picCtrlSrc.ScaleWidth / 2
cWy = picCtrlSrc.ScaleHeight / 2

' resize buttons to proper width, height
For i = 0 To 2
picBtn(i).Width = Wx
picBtn(i).Height = Wy
Next

For i = 0 To 1
picCtrl(i).Width = cWx
picCtrl(i).Height = cWy
Next

'blit some pictures onto buttons
BitBlt picBtn(1).hdc, 0, 0, Wx, Wy, picSkin.hdc, picBtn(1).Left, picBtn(1).Top, vbSrcCopy
BitBlt picBtn(0).hdc, 0, 0, Wx, Wy, picSkin.hdc, picBtn(0).Left, picBtn(0).Top, vbSrcCopy
BitBlt picBtn(2).hdc, 0, 0, Wx, Wy, picSkin.hdc, picBtn(2).Left, picBtn(2).Top, vbSrcCopy

BitBlt picScroll.hdc, 0, 0, picScroll.ScaleWidth, picScroll.ScaleHeight, picSkin.hdc, picScroll.Left, picScroll.Top, vbSrcCopy


BitBlt picCtrl(0).hdc, 0, 0, cWx, cWy, picSkin.hdc, picCtrl(0).Left, picCtrl(0).Top, vbSrcCopy
BitBlt picCtrl(1).hdc, 0, 0, cWx, cWy, picSkin.hdc, picCtrl(1).Left, picCtrl(1).Top, vbSrcCopy


TransparentBlt picBtn(0).hdc, 0, 0, Wx, Wy, picBtnSrc.hdc, 0, 0, Wx, Wy, vbGreen
TransparentBlt picBtn(1).hdc, 0, 0, Wx, Wy, picBtnSrc.hdc, 0, Wy, Wx, Wy, vbGreen 'stop
TransparentBlt picBtn(2).hdc, 0, 0, Wx, Wy, picBtnSrc.hdc, 0, Wy * 3, Wx, Wy, vbGreen 'open

TransparentBlt picCtrl(0).hdc, 0, 0, cWx, cWy, picCtrlSrc.hdc, 0, 0, cWx, cWy, vbGreen
TransparentBlt picCtrl(1).hdc, 0, 0, cWx, cWy, picCtrlSrc.hdc, 0, cWy, cWx, cWy, vbGreen

BitBlt picTitle.hdc, 0, 0, picTitle.ScaleWidth, picTitle.ScaleWidth, picSkin.hdc, picTitle.Left, picTitle.Top, vbSrcCopy
picTitle.Refresh

BitBlt picAlbum.hdc, 0, 0, picAlbum.ScaleWidth, picAlbum.ScaleWidth, picSkin.hdc, picAlbum.Left, picAlbum.Top, vbSrcCopy
picAlbum.Refresh



' setup the player state variables
isPlaying = False
isNewFile = True

frmSplash.lblStatus.Caption = "Loading Player Settings ..."
#If Not SkinDBUG Then
   Dim skin As String
   skin = GetSetting(App.EXEName, "Settings", "Skin")
     If Dir(skin) <> "" Then
      modsFireSkinParser.renderSkin skin
     Else
      frmDummy.mnuChangeSkin_Click
     End If
     
#End If

' debug mode
#If DBUG Then
   curFile = "D:\ksk new\misc\anime\mp3\thank you, love.mp3"
   picBtn_Click 0
   Unload frmFirePL
#End If



#If Not DBUG Then
   frmFirePL.Show
#End If

' show startup parameters
Debug.Print String(30, "-")
Debug.Print "FireAMP! - Starting UP!" & vbCrLf
Debug.Print Abs(Timer - t) & " Seconds for start up, yeah!"

tFile = App.path & "\FireAMP.Media.Buffer"
doRepeatPlaylist = CBool(GetSetting(App.EXEName, "Settings", "Playlist Repeat", True))
doRandomPlayback = CBool(GetSetting(App.EXEName, "Settings", "Playlist Random", False))
isMinus = CBool(GetSetting(App.EXEName, "Settings", "isMinus", False))
frmDummy.mnuRepeatPlaylist.Checked = doRepeatPlaylist
frmDummy.mnuRandomPlayback.Checked = doRandomPlayback
fraDisplay.Visible = False
fraVideo.Visible = False


drawFlag = True
currentPart = 0
p = 1
infoStr = ""
lstPl.ColumnHeaders.Item(1).Width = CInt(lstPl.Width / 1.3)
lstPl.ColumnHeaders.Item(2).Width = lstPl.Width - lstPl.ColumnHeaders.Item(1).Width

Unload frmSplash
'frmFirePL.Hide
sleepFactor = 50

End Sub

Private Sub lblCaption_Click()
PopupMenu frmDummy.mnuFireAMP
End Sub

Private Sub lblStatus_Click()
isMinus = Not isMinus
SaveSetting App.EXEName, "Settings", "isMinus", isMinus
End Sub


Private Sub lstPl_DblClick()
frmFirePL.lstPl.ListItems(lstPl.SelectedItem.Index).Selected = True
frmFirePL.lstPl_DblClick
End Sub

Private Sub lstPl_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
 PopupMenu frmDummy.mnuPlaylist
End If
End Sub

Private Sub picBarBack_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
 tmrPbr.Enabled = False
 tmrTime.Enabled = False
picBarFront.Left = -picBarFront.ScaleWidth + X
End Sub

Private Sub picBarBack_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If Not (FireAMP_Pos Is Nothing) Then
On Error Resume Next
FireAMP_Pos.CurrentPosition = Abs(getBarPosition(picBarFront, picBarBack, FireAMP_Pos.Duration))
tmrPbr.Enabled = True
tmrTime.Enabled = True
End If
End Sub

Private Sub picBarFront_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
mx = X
PauseClip
End Sub

'seek bar
Public Sub picBarFront_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)


If Button = 1 And Not (FireAMP_Pos Is Nothing) And drawFlag Then
 tmrPbr.Enabled = False
 tmrTime.Enabled = False
 
 If picBarFront.Left < 0 Then
  picBarFront.Left = 0
  drawFlag = False
    
 ElseIf picBarFront.Left + picBarFront.ScaleWidth >= picBarBack.ScaleWidth Then
  picBarFront.Left = picBarBack.ScaleWidth - picBarFront.ScaleWidth
  drawFlag = False
   
 End If
 
 
  picBarFront.Move picBarFront.Left + X - mx
  lblStatus.Caption = "Seeking... " & convertToStdTime(getBarPosition(picBarFront, picBarBack, FireAMP_Pos.Duration))
  
 End If
End Sub

Public Sub picBarFront_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If Not (FireAMP_Pos Is Nothing) Then
On Error Resume Next
FireAMP_Pos.CurrentPosition = (getBarPosition(picBarFront, picBarBack, FireAMP_Pos.Duration))
ResumeClip
tmrPbr.Enabled = True
tmrTime.Enabled = True
drawFlag = True
End If
End Sub

' play, stop and open
Public Sub picBtn_Click(Index As Integer)
On Local Error Resume Next
Static isID3Filled As Boolean

Dim LST As ListItem, e As ErrStruct
Dim Count As Integer

Select Case Index
Case 0:
Debug.Print "Debug: picBtn_Click 0 (Play)"
If isNewFile Then
If curFile = "" Then Exit Sub
  picBtn_Click 1
  isNewFile = False
  isPlaying = True
lblTitle.Caption = "Opening Clip..."
Pos = 1
Dim ext As String
ext = getFileExtensionFromPath(curFile)
 If Not isVideoFile(curFile) Then
 ' playing an audio file
      setupVisualization
      fraDisplay.Visible = True
'picPlaylist.Visible = False
      'copy to buffer
                If Not Fsys.FileExists(curFile) Then
                    e.errNum = 12
                    e.errShortDesc = "File not found"
                    e.errLongDesc = "File not found : " & curFile
                    logError e
                    Exit Sub
                End If
 
    FileCopy curFile, tFile
    PlayClip tFile
    
picTitle.Visible = True
picAlbum.Visible = True
picVisual.Visible = True
lblInfo.Visible = True

Else
 'video files are played directly. takes too long to copy and access
        PlayClip curFile
    fraDisplay.Visible = False
picTitle.Visible = True
fraVideo.Visible = True
'picPlaylist.Visible = False
End If
tmrPbr.Enabled = True
tmrTime.Enabled = True
picBtn_MouseUp 0, 1, 0, 0, 0
infoStr = "[Playing] " & getFileTitleFromPath(curFile) & Chr(0) & " FireAMP! "


lblStatus.Visible = True
picBarBack.Visible = True

If getFileExtensionFromPath(curFile) = "mp3" Then
    frmDummy.mnuGoogleSearch.Enabled = True
    Dim tempTag As tagID3_1x
    tempTag = readID3_1x(curFile)
    
    lblTitle.Caption = Trim(tempTag.Artist) & " " & Chr(CLng(sepChar)) & " " & Trim(tempTag.Title)
    lblAlbum.Caption = Trim(tempTag.Album)
    infoStr1 = Space(8) & lblTitle.Caption
    
    
    Dim bitrateInfo As mp3HeaderData
    bitrateInfo = ReadMP3Header(curFile)
    
    infoStr = "Clip: " & tempTag.Title
    infoStr = infoStr & Chr(0) & "Artist: " & tempTag.Artist
    infoStr = infoStr & Chr(0) & "Album: " & tempTag.Album
    infoStr = infoStr & Chr(0) & "Playing at " & bitrateInfo.BitRate & "Kbps"
    
    If Not isID3Filled Then
       If Val(GetSetting(App.EXEName, "Settings", "Alerts", "1")) = 0 Then frmAlert.Show vbModal
    End If
Else
    lblTitle.Caption = getFileTitleFromPath(Trim(curFile))
    lblAlbum.Caption = UCase(getFileExtensionFromPath(Trim(curFile))) & " Media"
    infoStr1 = Space(8) & lblTitle.Caption
End If

For Count = 1 To frmFirePL.lstPl.ListItems.Count
    If frmFirePL.lstPl.ListItems.Item(Count).Bold Then frmFirePL.lstPl.ListItems.Item(Count).Bold = False
    frmFirePL.lstPl.ListItems.Item(Count).ForeColor = plColor
Next Count
 
    Set LST = frmFirePL.lstPl.ListItems.Item(frmFirePL.lstPl.SelectedItem.Index)
    LST.SubItems(1) = convertToStdTime(FireAMP_Pos.Duration)
    frmFirePL.lstPl.SelectedItem.Bold = True
    frmFirePL.lstPl.SelectedItem.ForeColor = plCurrent
    LST.EnsureVisible
    LST.ForeColor = plCurrent
    BitBlt picScroll.hdc, 0, 0, picScroll.ScaleWidth, picScroll.ScaleWidth, picSkin.hdc, picScroll.Left, picScroll.Top, vbSrcCopy
    
    isPlaying = True
    isPlaying = True
      picBtn_MouseUp 0, 1, 0, 0, 0
      
   delay = 0
   X = 0
        
   tmrInfo.Enabled = True
   
   Debug.Print , "File: ", curFile
   Debug.Print , "Type: ", getFileExtensionFromPath(curFile)
   Debug.Print "Status: ", IIf(isPlaying, "Playing", "Paused")
   Debug.Print "End Debug"
   Exit Sub
End If

Dim Control As Object
If isPlaying Then
 PauseClip
 lblCaption.Caption = "FireAMP! [Paused]"
  tmrInfo.Enabled = False
 lblInfo.Caption = "Paused"
 If isMute Then lblCaption.Caption = lblCaption.Caption & " [Mute]"

Else
 ResumeClip
 lblCaption.Caption = "FireAMP!"
  tmrInfo.Enabled = True
  If isMute Then lblCaption.Caption = lblCaption.Caption & " [Mute]"

End If


   Debug.Print , "File: ", curFile
   Debug.Print , "Type: ", getFileExtensionFromPath(curFile)
   Debug.Print "Status: ", IIf(isPlaying, "Playing", "Paused")
   Debug.Print " End Debug"
isPlaying = Not isPlaying
 picBtn_MouseUp 0, 1, 0, 0, 0

Case 1:

StopClip
DoStop
isNewFile = True
isPlaying = False
picBtn_MouseUp 0, 1, 0, 0, 0
If Not tFile = "" Then Fsys.DeleteFile tFile, True
Set oPlugIn = Nothing
lblStatus.Caption = "00:00 [00:00]"

lblStatus.Visible = False
picBarBack.Visible = False
fraVideo.Visible = False
picVisual.Visible = False
picTitle.Visible = False

picAlbum.Visible = False
'picPlaylist.Visible = True
tmrPbr.Enabled = False
tmrTime.Enabled = False
tmrInfo.Enabled = False

infoStr = ""
lblInfo.Visible = False
frmFirePL.lstPl.ListItems(playingIndex).Bold = False
Caption = "FireAMP!"


Case 2:
cd1.FileName = ""
cd1.ShowOpen
curFile = cd1.FileName
If curFile <> "" Then
StopClip
frmFirePL.lstPaths.AddItem curFile

Set LST = frmFirePL.lstPl.ListItems.Add(, , getFileTitleFromPath(curFile))
Dim Temp_Player As New FilgraphManager
Dim Temp_Pos As IMediaPosition
On Error GoTo e
Temp_Player.RenderFile curFile
Set Temp_Pos = Temp_Player
LST.SubItems(1) = convertToStdTime(Temp_Pos.Duration)

Set Temp_Pos = Nothing
Set Temp_Player = Nothing
 picBtn_Click 0
End If


End Select
Exit Sub
e:
LST.SubItems(1) = "??:??"
LST.ForeColor = vbRed
End Sub

Private Sub picBtn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
Dim offSetY As Integer
Select Case Index
Case 0: 'play
If isPlaying Then
offSetY = Wy * 2
Else
offSetY = 0
End If

Case 1: 'stop
offSetY = Wy
Case 2: 'open
offSetY = Wy * 3
End Select

TransparentBlt picBtn(Index).hdc, 0, 0, Wx, Wy, picBtnSrc.hdc, Wx, offSetY, Wx, Wy, vbGreen
picBtn(Index).Refresh
End Sub

Private Sub picBtn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
Dim offSetY As Integer
Select Case Index
Case 0: 'play
If isPlaying Then
offSetY = Wy * 2
Else
offSetY = 0
End If

Case 1: 'stop
offSetY = Wy
Case 2: 'open
offSetY = Wy * 3
End Select

TransparentBlt picBtn(Index).hdc, 0, 0, Wx, Wy, picBtnSrc.hdc, 0, offSetY, Wx, Wy, vbGreen
picBtn(Index).Refresh

End Sub

Private Sub picCtrl_Click(Index As Integer)
Select Case Index
Case 0:
 Me.WindowState = vbMinimized
frmFirePL.Visible = False
Case 1:
picBtn_Click 1
Form_Unload 0
End Select
End Sub

Private Sub picCtrl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case Index
Case 0
TransparentBlt picCtrl(0).hdc, 0, 0, cWx, cWy, picCtrlSrc.hdc, cWx, 0, cWx, cWy, vbGreen
Case 1
TransparentBlt picCtrl(1).hdc, 0, 0, cWx, cWy, picCtrlSrc.hdc, cWx, cWy, cWx, cWy, vbGreen
End Select
picCtrl(Index).Refresh
End Sub

Private Sub picCtrl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case Index
Case 0
BitBlt picCtrl(0).hdc, 0, 0, cWx, cWy, picCtrlSrc.hdc, 0, 0, vbSrcCopy
Case 1
BitBlt picCtrl(1).hdc, 0, 0, cWx, cWy, picCtrlSrc.hdc, 0, cWy, vbSrcCopy
End Select
picCtrl(Index).Refresh
End Sub


' keyboard short cuts
Public Sub picSkin_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case Chr(KeyCode)
 Case "X", "x": ' goodbye
Form_Unload 0
 
 Case "F", "f": ' full screen video
 Dim ext As String
 ext = getFileExtensionFromPath(curFile)
 If ext = "mpg" Or ext = "mpeg" Or ext = "dat" Or ext = "wmv" Or ext = "rm" Or ext = "rmvb" Or ext = "mov" Or ext = "avi" Then
 FireAMP_VideoWin.HideCursor True
  refreshVideo frmFullScreen.Frame1
 frmFullScreen.Show
 Me.Hide
 End If
Case " ": ' play
picBtn_Click 0
Case "S", "s": ' stop
picBtn_Click 1
Case "O", "o": ' open
picBtn_Click 2
Case "P", "p": ' show playlist

Case "G", "g": ' volume up
changeVolume True

Case "H", "h": ' volume down
 changeVolume False

Case "M", "m": 'mute
 If isMute Then
  FireAMP_Vol.Volume = 0
lblCaption.Caption = "FireAMP !"
Else
 FireAMP_Vol.Volume = -10000
lblCaption.Caption = "FireAMP ! [Mute]"
End If
isMute = Not isMute

Case "N", "n":
 picCtrl_Click (0)
Case "C", "c"
 PlayVideoCD
Case "r", "R":
 frmExtVis.Show
End Select


End Sub

Private Sub picSkin_LostFocus()
If Not FireAMP_Pos Is Nothing Then tmrVisUpdate.Enabled = True
End Sub

Private Sub picSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 Then
      ReleaseCapture
      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
SetFocus
If Button = 2 Then
  frmDummy.PopupMenu frmDummy.mnuVis
End If

End Sub


Private Sub changeVolume(UPorDOWN As Boolean)
If Not (FireAMP_Pos Is Nothing) Then

If UPorDOWN Then  ' vol. up
currentVolume = currentVolume - 500
If currentVolume < -5000 Then currentVolume = -5000
FireAMP_Vol.Volume = currentVolume
Else ' vol. down
currentVolume = currentVolume + 500
If currentVolume > 0 Then currentVolume = 0
FireAMP_Vol.Volume = currentVolume
End If
lblStatus.Caption = "Vol: " & (5000 + currentVolume) / 100 * 2 & "%"
End If

End Sub


Private Sub picVisual_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If Not FireAMP_Pos Is Nothing Then
 tmrVisUpdate.Enabled = True
 ReleaseCapture
 End If
End Sub

Private Sub tmrInfo_Timer()
On Error Resume Next
Static tColor As Long

If lblTitle.Width < picTitle.Width Then
picTitle.Cls
BitBlt picTitle.hdc, 0, 0, picTitle.ScaleWidth, picTitle.ScaleWidth, picSkin.hdc, picTitle.Left, picTitle.Top, vbSrcCopy
picTitle.Print lblTitle.Caption
Else

picTitle.Cls
BitBlt picTitle.hdc, 0, 0, picTitle.ScaleWidth, picTitle.ScaleWidth, picSkin.hdc, picTitle.Left, picTitle.Top, vbSrcCopy
 

picScrollBuffer.BackColor = tColor

picScrollBuffer.Cls
picScrollBuffer.Width = lblTitle.Width * 1.2

picScrollBuffer.Print lblTitle.Caption

If (picScrollBuffer.ScaleWidth * 1.2 + X) < 0 Then X = 0: delay = 0
        TransparentBlt picTitle.hdc, X, 0, picScrollBuffer.ScaleWidth, picScrollBuffer.ScaleHeight, picScrollBuffer.hdc, 0, 0, picScrollBuffer.ScaleWidth, picScrollBuffer.ScaleHeight, tColor
        TransparentBlt picTitle.hdc, picScrollBuffer.ScaleWidth * 1.2 + X, 0, picScrollBuffer.ScaleWidth, picScrollBuffer.ScaleHeight, picScrollBuffer.hdc, 0, 0, picScrollBuffer.ScaleWidth, picScrollBuffer.ScaleHeight, tColor
        picTitle.Refresh
If delay > 100 Then X = X - 1

delay = delay + 1
End If

Me.Caption = lblTitle.Caption & " " & lblStatus.Caption
Pos = Pos + 1
'Debug.Print CInt(Pos)

If CInt(Pos) > Len(infoStr) Then Pos = 1

InfoParts = Split(infoStr, Chr(0))
lblInfo.Caption = Mid(InfoParts(currentPart), 1, p)

If Not rFlag Then
 p = p + 0.5
Else
 p = p - 0.5
End If

If CInt(p) > Len(InfoParts(currentPart)) * 1.5 Then rFlag = True

If p = 0 Then
 currentPart = currentPart + 1
 p = 1
 rFlag = False
End If

If currentPart > UBound(InfoParts) Then currentPart = 0
End Sub

Private Sub tmrPbr_Timer()

'If (Not (FireAMP_Pos Is Nothing)) Then
'If FireAMP_Pos.Duration = FireAMP_Pos.CurrentPosition And frmFirePL.lstPaths.ListIndex < frmFirePL.lstPaths.ListCount - 1 Then
'
'lblTitle.Caption = "Changing Track..."
'lblAlbum.Caption = ""
'
'picBtn_Click 1
'curFile = frmFirePL.lstPaths.List(frmFirePL.lstPaths.ListIndex + 1)
'frmFirePL.lstPaths.ListIndex = frmFirePL.lstPaths.ListIndex + 1
'On Local Error Resume Next
'frmFirePL.lstPL.ListItems(frmFirePL.lstPaths.ListIndex + 1).Selected = True
'picBtn_Click 0
'
'End If
'On Local Error Resume Next
'picBarFront.Left = (picBarBack.ScaleWidth - picBarFront.Width) * (FireAMP_Pos.CurrentPosition / FireAMP_Pos.Duration)
'
'Else
' fraVideo.Visible = False
' picVisual.Visible = False
' DoStop
' On Error Resume Next
' frmFirePL.lstPaths.ListIndex = 0
' tmrPbr_Timer
'End If
Dim currentIndex As Integer
If FireAMP_Pos Is Nothing Then
tmrTime.Enabled = False
 tmrPbr.Enabled = False
 Exit Sub
End If
On Error Resume Next
If FireAMP_Pos.Duration = FireAMP_Pos.CurrentPosition Then
       If frmFirePL.lstPaths.ListCount = 0 Then
          tmrPbr.Enabled = False
       Else
       If Not doRandomPlayback Then
           If frmFirePL.lstPaths.ListIndex < frmFirePL.lstPaths.ListCount - 1 Then
              frmFirePL.lstPaths.ListIndex = frmFirePL.lstPaths.ListIndex + 1
           Else
               If doRepeatPlaylist Then
                 frmFirePL.lstPaths.ListIndex = 0
               Else
                 picBtn_Click 0
               End If
           End If
       Else
       Randomize Timer
       frmFirePL.lstPaths.ListIndex = Rnd * frmFirePL.lstPaths.ListCount + 1
       
            
       End If
              currentIndex = frmFirePL.lstPaths.ListIndex + 1
              
              
       frmFirePL.lstPl.ListItems(currentIndex).Selected = True
       curFile = frmFirePL.lstPaths.List(frmFirePL.lstPaths.ListIndex)
            
        picBtn_Click 1
        picBtn_Click 0
        
'        If frmMediaLib.Visible Then
'        frmMediaLib.lstMediaLib.ListItems(frmMediaLib.lstMediaLib.SelectedItem.Index + 1).Selected = True
'       End If

       End If

Else
  updateBar picBarFront, picBarBack, FireAMP_Pos.Duration, FireAMP_Pos.CurrentPosition
End If

 
End Sub

Private Sub tmrTime_Timer()
On Error Resume Next
If Not FireAMP_Pos Is Nothing Then
If Not isMinus Then
lblStatus.Caption = " " & convertToStdTime(FireAMP_Pos.CurrentPosition) & " [" & convertToStdTime(FireAMP_Pos.Duration) & "]"
Else
lblStatus.Caption = "-" & convertToStdTime(FireAMP_Pos.Duration - FireAMP_Pos.CurrentPosition) & " [" & convertToStdTime(FireAMP_Pos.Duration) & "]"
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim p As Object
For Each p In Forms
Set p = Nothing
Next p
 Set Fsys = Nothing ' release FSys
    If DevHandle <> 0 Then
        Call DoStop
End If
End
End Sub

' Initialize visualization
Public Sub setupVisualization()
initWaveIn

    Call DoReverse
    
    ScopeBuff.Width = picVisual.Width
    ScopeBuff.Height = picVisual.Height
    
    ScopeBuff.BackColor = picVisual.BackColor
        
On Error GoTo No_PI
Set oPlugIn = CreateObject(GetSetting(App.EXEName, "Visualization", "Object"))
frmDummy.mnuConfig.Enabled = CBool(GetSetting(App.EXEName, "Visualization", "Config"))
tmrVisUpdate.Enabled = True
picVisual.ToolTipText = GetSetting(App.EXEName, "Visualization", "Name")

Exit Sub

No_PI:
Debug.Print "Falied to load plugin"
Err.Clear
'frmDummy.mnuLoadPi_Click
DoStop
 'setupVisualization
End Sub


Sub PlayVideoCD()
curFile = "E:\mpegav\avseq01.dat"
If Not Fsys.FileExists(curFile) Then
 Dim e As ErrStruct
 e.errNum = 3
 e.errShortDesc = "Invalid CD in drive"
 e.errLongDesc = "FireAMP! tried to play a video CD but it appears that the CD in the drive is not a videoCD" & vbCrLf & "Replace with a videoCD and try again"
 logError e
 Exit Sub
End If
picBtn_Click 0
End Sub

Private Sub tmrVisUpdate_Timer()
On Error GoTo e
ScopeBuff.Cls

    Static Wave As WAVEHDR


    Static InData(0 To NUMSAMPLES - 1) As Integer      ' wave-in data
    
        'lpdata requires the address of an array to fill up data with
            Wave.lpData = VarPtr(InData(0))
        'the buffer length
            Wave.dwBufferLength = NUMSAMPLES
        ' ???
            Wave.dwFlags = 0
         'prepare device for input
            Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
            Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))

            ' if the following statement is removed, the vis. will be a lot faster (avs style)
            ' but uses up 100% of cpu!
            ' this is why i hate avs
            Sleep sleepFactor ' give device a breather

            ' the following loop is quite useless, but anyway...
            Do
                'Just wait for the blocks to be done or the device to close
            Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0


            If DevHandle = 0 Then Exit Sub 'Cut out if the device is closed

            Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
                        
' now call the drawVis method of the pi
oPlugIn.drawVis ScopeBuff.hdc, InData, ScopeBuff.ScaleHeight, ScopeBuff.ScaleWidth
picVisual.Picture = ScopeBuff.Image

If frmExtVis.Visible Then
StretchBlt frmExtVis.hdc, 0, 0, frmExtVis.ScaleWidth, frmExtVis.ScaleHeight, frmFireMain.ScopeBuff.hdc, 0, 0, frmFireMain.ScopeBuff.ScaleWidth, frmFireMain.ScopeBuff.ScaleHeight, vbSrcCopy
frmExtVis.Refresh
End If

Exit Sub
e:
picVisual.Cls
picVisual.ForeColor = RGB(192, 129, 129)
picVisual.Print "Error #" & Err.Number

picVisual.ForeColor = vbRed
picVisual.Print "An error has occured in this PlugIn! [" + Err.Description + "]"

picVisual.ForeColor = RGB(192, 96, 96)
picVisual.Print "Try Reloading or choose another PlugIn"
DoStop
 tmrVisUpdate.Enabled = False
End Sub
