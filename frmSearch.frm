VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Playlist"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Match Case"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Restart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   2040
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   2760
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   -15
      Top             =   1920
      Width           =   7095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search for:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0 Items found"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   5055
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LST As ListItem, searched As Boolean, j As Integer
Dim matchCase As Boolean

Private Sub Check1_Click()
matchCase = Check1.Value
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Not searched Then
Dim i As Integer
For i = 1 To frmFirePL.lstPl.ListItems.Count
If InStr(1, frmFirePL.lstPl.ListItems.Item(i).Text, Text1.Text, IIf(matchCase, vbBinaryCompare, vbTextCompare)) Then
frmFirePL.lstPl.ListItems.Item(i).ForeColor = vbWhite
List1.AddItem i
End If
Next
Label1.Caption = List1.ListCount & " Items found"
Label1.Caption = Label1.Caption & ". Press the play button to browse."
Else
frmFirePL.lstPl.ListItems.Item(Val(List1.List(j))).EnsureVisible
frmFirePL.lstPl.ListItems.Item(Val(List1.List(j))).ForeColor = vbGreen
If j > 0 Then frmFirePL.lstPl.ListItems.Item(Val(List1.List(j - 1))).ForeColor = vbWhite
j = j + 1
If j > List1.ListCount - 1 Then j = 0
End If
searched = True

End Sub

Private Sub Command2_Click()
Unload Me
frmSearch.Show
End Sub

Private Sub Form_Load()
searched = False
j = 0
Dim Count As Integer
For Count = 1 To frmFirePL.lstPl.ListItems.Count
frmFirePL.lstPl.ListItems.Item(Count).ForeColor = plColor
Next Count
Check1.Value = IIf(matchCase, 1, 0)
End Sub

