VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5400
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1425
      Left            =   4080
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1425
      ScaleWidth      =   1380
      TabIndex        =   6
      Top             =   0
      Width           =   1380
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4560
      Top             =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   -240
      Top             =   3360
      Width           =   5655
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "(You'll looove Orange Soda)"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   8
      Top             =   2760
      Width           =   1980
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Powered by Orange Soda"
      ForeColor       =   &H000040C0&
      Height          =   195
      Left            =   1320
      TabIndex        =   7
      Top             =   2400
      Width           =   1830
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Created by K.Sai Krishna and K.V.Rohit"
      ForeColor       =   &H00954A00&
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   1560
      Width           =   2805
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kick Winamp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2055
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "(Version Info)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FireAMP!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006060C0&
      Height          =   555
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' the about form
'

' last updated: 2006 May 05, Humanoid

Dim i As Integer
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim abt(0 To 1) As String
abt(0) = "The Media Player, Rewinded"
abt(1) = "Music::Genre->All"
Label3.Caption = "(Version " & App.Major & "." & App.Minor & ", Build " & App.Revision & ")"
Label5.Caption = "Created by K.Sai Krishna and K.V.Rohit (and perhaps Panda)"
Randomize Time
Label4.Caption = abt(Rnd * 1)
i = 255
End Sub

Private Sub Timer1_Timer()
Label4.ForeColor = RGB(i, i, i)
Label7.ForeColor = RGB(i * 2, i * 2, i * 2)
i = i - 10
If i < 50 Then Timer1.Enabled = False
End Sub

