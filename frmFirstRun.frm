VERSION 5.00
Begin VB.Form frmFirstRun 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FireAMP! 1st Run"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frmFirstRun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3840
      Top             =   5400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Shape Shape5 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   -1560
         Top             =   4320
         Width           =   9975
      End
      Begin VB.Image Image1 
         Height          =   2385
         Left            =   1800
         Picture         =   "frmFirstRun.frx":000C
         Top             =   1320
         Width           =   4410
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   855
         Left            =   2160
         TabIndex        =   1
         Top             =   1440
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   360
      TabIndex        =   2
      Top             =   0
      Width           =   7815
      Begin VB.PictureBox picHolder2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   5760
         ScaleHeight     =   495
         ScaleWidth      =   1935
         TabIndex        =   9
         Top             =   4440
         Width           =   1935
         Begin VB.CommandButton Command5 
            Caption         =   "&Done"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.PictureBox picHolder1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   3000
         ScaleHeight     =   2295
         ScaleWidth      =   2055
         TabIndex        =   4
         Top             =   1440
         Width           =   2055
         Begin VB.CommandButton Command1 
            Caption         =   "&Set up Stereo Mix"
            Height          =   375
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1935
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Choose a Visualization"
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Top             =   600
            Width           =   1935
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Se&lect a Skin"
            Height          =   375
            Left            =   0
            TabIndex        =   6
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton Command4 
            Caption         =   "&Edit Preferences"
            Height          =   375
            Left            =   0
            TabIndex        =   5
            Top             =   1800
            Width           =   1935
         End
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   -120
         Top             =   4320
         Width           =   9495
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2760
         Shape           =   3  'Circle
         Top             =   3360
         Width           =   135
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2760
         Shape           =   3  'Circle
         Top             =   2760
         Width           =   135
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2760
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   135
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   2760
         Shape           =   3  'Circle
         Top             =   1560
         Width           =   135
      End
      Begin VB.Line Line5 
         X1              =   2040
         X2              =   2760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line4 
         X1              =   2040
         X2              =   2760
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line3 
         X1              =   2040
         X2              =   2760
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line2 
         X1              =   2040
         X2              =   2760
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   2040
         Y1              =   1320
         Y2              =   3480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To do:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   960
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmFirstRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim delay As Integer
Dim StereoMixIsDone As Boolean, VisualizationIsSelected As Boolean
Dim SkinIsChosen As Boolean
Private Sub Command1_Click()
frmDummy.mnuAdjustLevels_Click
StereoMixIsDone = True
End Sub

Private Sub Command2_Click()
frmDummy.mnuLoadPi_Click
VisualizationIsSelected = True
End Sub

Private Sub Command3_Click()
frmDummy.cd1.FileName = ""
frmDummy.cd1.Filter = "FireAMP! Skins (*.cfs)|*.cfs"
frmDummy.cd1.ShowOpen
If frmDummy.cd1.FileName <> "" Then
 SaveSetting App.EXEName, "Settings", "Skin", frmDummy.cd1.FileName
 SkinIsChosen = True
End If

End Sub

Private Sub Command4_Click()
frmOptions.Show vbModal
End Sub

Private Sub Command5_Click()
If VisualizationIsSelected And StereoMixIsDone And SkinIsChosen Then
frmFireMain.Show
frmFirePL.Show
Unload Me
Fsys.CreateTextFile App.path & "\FireAMP.Options"
Else
Dim e As ErrStruct
e.errNum = 0
e.errShortDesc = "Missing requisite conditions"
e.errLongDesc = "Please Select the Stereo Mix Option and Choose a Visualization"
logError e
End If

End Sub

Private Sub Timer1_Timer()
delay = delay + 1
If delay > 50 Then
Frame1.Left = Frame1.Left - 50
Frame2.Left = Frame1.Left + Frame1.Width
End If

If Frame2.Left < 200 Then Timer1.Enabled = False

End Sub
