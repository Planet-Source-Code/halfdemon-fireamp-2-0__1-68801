VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Replace File Name"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "frmRename.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "From:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "To:"
      Height          =   195
      Left            =   1680
      TabIndex        =   6
      Top             =   3120
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
File1.Path = Dir1.Path
Dim i As Integer, oldName As String, newName As String
Dim Fsys As New FileSystemObject

For i = 0 To File1.ListCount - 1
oldName = File1.List(i)
newName = Replace(oldName, Text1.Text, Text2.Text, , , vbBinaryCompare)
'FileCopy File1.Path & oldfile, File1.Path & newfile
Debug.Print oldName, newName
Next

Set Fsys = Nothing
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive & "\"
End Sub

