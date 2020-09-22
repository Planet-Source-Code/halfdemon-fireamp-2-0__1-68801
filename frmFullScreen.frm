VERSION 5.00
Begin VB.Form frmFullScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   520
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   887
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   3840
      Width           =   13095
      Begin VB.PictureBox picBar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   263
         TabIndex        =   2
         Top             =   120
         Width           =   3975
         Begin VB.Label picBarFront 
            BackColor       =   &H005C5CCD&
            ForeColor       =   &H005C5CCD&
            Height          =   255
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblTitle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   375
         Left            =   8040
         TabIndex        =   6
         Top             =   120
         Width           =   5055
      End
      Begin VB.Label lblVideo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FireAMP! Video"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   330
         Left            =   5760
         TabIndex        =   5
         Top             =   70
         Width           =   2130
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "00:00 [00:00]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H005C5CCD&
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3360
      Top             =   4440
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   5535
   End
End
Attribute VB_Name = "frmFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s As String
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If LCase(Chr(KeyCode)) = "f" Then
FireAMP_VideoWin.HideCursor False
refreshVideo frmFireMain.fraVideo
 frmFireMain.fraVideo.Visible = True
 frmFireMain.Show

Unload Me
Else
 frmFireMain.picSkin_KeyDown KeyCode, Shift
End If
End Sub

Private Sub Form_Load()
Frame1.Top = 0
Frame1.Left = 0
Frame1.Height = Screen.Height / Screen.TwipsPerPixelY - Frame2.Height
Frame1.Width = Screen.Width / Screen.TwipsPerPixelX

Frame2.Left = 0
Frame2.Top = Frame1.Top + Frame1.Height
lblTitle.Caption = "[ Playing: " & getFileTitleFromPath(curFile) & " ]"
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
picBarFront.Width = (picBar.ScaleWidth) * (FireAMP_Pos.CurrentPosition / FireAMP_Pos.Duration)
lblTime.Caption = convertToStdTime(FireAMP_Pos.CurrentPosition) & " [" & convertToStdTime(FireAMP_Pos.Duration) & "]"

If s <> curFile Then
 lblTitle.Caption = "[ Playing: " & getFileTitleFromPath(curFile) & " ]"
 s = curFile
End If

ext = getFileExtensionFromPath(curFile)
 If Not (ext = "mpg" Or ext = "mpeg" Or ext = "dat" Or ext = "wmv" Or ext = "rm" Or ext = "rmvb" Or ext = "mov" Or ext = "avi") And Not IsNull(FireAMP_Pos) Then
 Unload Me
 frmFireMain.Show
 frmFireMain.picVisual.Visible = True
End If
End Sub
