VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ID3 Missing"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000010&
      Caption         =   "Never show this alert again"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAlert.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "to launch the ID3 Editor."
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This Clip has missing ID3 Tag data."
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   1080
      MouseIcon       =   "frmAlert.frx":044E
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   -240
      Top             =   1560
      Width           =   7935
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' the FireAMP! alerter
'

' last updated: 2006 Aug 27, Humanoid

Private Sub Check1_Click()
SaveSetting App.EXEName, "Settings", "Alerts", CStr(Abs(Check1.Value))
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub lblEdit_Click()
'frmTagEditor.FileName = curFile
'Unload Me
'frmTagEditor.Show vbModal
End Sub

Private Sub picAlert_Click()
Unload Me
End Sub

