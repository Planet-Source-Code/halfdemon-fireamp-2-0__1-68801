VERSION 5.00
Begin VB.Form frmFireTrap 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FireAMP ErrorTrap"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Details"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   -15
      Top             =   1440
      Width           =   4995
   End
   Begin VB.Label lblNum 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   1980
   End
   Begin VB.Label lblReason 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmFireTrap.frx":0000
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblError 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "frmFireTrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim State As Boolean
Dim t As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
State = Not State
If State Then
t = lblReason.Top
Height = Height + lblReason.Height + 100
Command1.Top = Command1.Top + lblReason.Height + 100
Command2.Top = Command1.Top
Else
lblReason.Top = t
Height = Height - lblReason.Height - 100
Command1.Top = Command1.Top - lblReason.Height - 100
Command2.Top = Command1.Top
End If
Shape1.Top = Command1.Top - 100
End Sub

Private Sub Form_Load()
MessageBeep vbError ' beep away...
State = False
End Sub
