VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFrontEND 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RegSvr32 Front End"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   360
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Dynamic Link Libraries(*.dll)|*.dll"
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Frame fraSelect 
      Caption         =   "Select File"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3360
         ScaleHeight     =   495
         ScaleWidth      =   1815
         TabIndex        =   3
         Top             =   720
         Width           =   1815
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse..."
            Height          =   375
            Left            =   600
            TabIndex        =   4
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmFrontEND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdBrowse_Click()
    CD.FileName = ""
    CD.ShowOpen
    If CD.FileName = "" Then Exit Sub
    txtFile.Text = CD.FileName
End Sub

Private Sub cmdRegister_Click()
    ShellExecute Me.hwnd, "Open", "regsvr32", """" & txtFile.Text & """", "", 0
End Sub
