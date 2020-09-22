VERSION 5.00
Begin VB.Form frmAddClips 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Folder of Music"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Label lblProgress 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   4095
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdv 
      Caption         =   "&View Files"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Done"
      Height          =   375
      Left            =   9240
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Include Subfolders"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   5400
      TabIndex        =   3
      Top             =   600
      Width           =   4935
   End
   Begin VB.ComboBox cboTypes 
      Height          =   315
      ItemData        =   "frmAddClips_new.frx":0000
      Left            =   5640
      List            =   "frmAddClips_new.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   -15
      Top             =   3720
      Width           =   10935
   End
End
Attribute VB_Name = "frmAddClips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' the add folder dialog
'

' last updated: 2006 May 05, Humanoid

' toggle for Advanced view
Private isAdvanced As Boolean
Private addSubfolders As Boolean

Private Sub cboTypes_Click()
Dim Exts(0 To 12) As String
' media types
Exts(0) = "*.mp3;*.mp2;*.mp1;*.mid;*.rmi;*.wav;*.rm;*.rmvb;*.mov;*.mpg;*.mpg;*.mpe;*.wma;*.wmv"
Exts(1) = "*.mp3;*.mp2;*.mp1;*.mid;*.rmi;*.wav;*.wma"
Exts(2) = "*.mpg;*.mpg;*.mpe;*.wmv"
Exts(3) = "*.mid;*.rmi"
Exts(4) = "*.mp3;*.mp2;*.mp1"
Exts(5) = "*.wma"
Exts(6) = "*.rm;*.rmvb"
Exts(7) = "*.mov"
Exts(8) = "*.mpg;*.mpe;*.mpeg"
Exts(9) = "*.wmv"
Exts(10) = "*.wav"
Exts(11) = "*.*"

             
File1.Pattern = CStr(Exts(cboTypes.ListIndex))
             
End Sub

Private Sub Check1_Click()
addSubfolders = Check1.Value
End Sub

Private Sub cmdAdd_Click()
Dim file As file
Frame2.Visible = True
cmdAdd.Enabled = False
cmdAdv.Enabled = False
modCommon.stopScan = False
If addSubfolders Then
scanFolder Dir1.path, frmFirePL.lstPaths, frmFirePL.lstPl, , , , lblProgress
Else
For Each file In Fsys.GetFolder(Dir1.path).Files
 If isMediaFile(file.path) Then
  frmFirePL.lstPaths.AddItem file.path
  frmFirePL.lstPl.ListItems.Add , , getFileTitleFromPath(file.Name)
 End If
Next
End If

'Dim i As Integer
'For i = 1 To frmFirePL.lstPl.ListItems.Count
'frmFireMain.lstPl.ListItems.Add , , frmFirePL.lstPl.ListItems(i).Text
'Next

Unload Me
End Sub

Private Sub cmdAdv_Click()
' toggle
isAdvanced = Not isAdvanced
 If isAdvanced Then
  Width = Width + 5265
  cmdAdv.Caption = "Hide Files"
 Else
  Width = Width - 5265
  cmdAdv.Caption = "View Files"
 
 End If
End Sub

Private Sub cmdCancel_Click()
modCommon.stopScan = True
Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
File1.path = Dir1.path
oldPath = Dir1.path
End Sub

Private Sub Drive1_Change()
On Error GoTo 1
Dir1.path = Drive1.Drive & "\"
Exit Sub

1:
 Dim e As ErrStruct
 e.errNum = Err.Number
 e.errShortDesc = Err.Description
 e.errLongDesc = Err.Description
 
 logError e
End Sub

Private Sub File1_DblClick()
If Right(File1.path, 1) = "\" Then
frmFirePL.lstPaths.AddItem File1.path & File1.FileName
Else
frmFirePL.lstPaths.AddItem File1.path & "\" & File1.FileName
End If
frmFirePL.lstPl.ListItems.Add , , getFileTitleFromPath(File1.FileName)
End Sub

Private Sub Form_Load()
isAdvanced = False
cboTypes.AddItem "All known media files"
cboTypes.AddItem "All Audio files"
cboTypes.AddItem "All Video files"
cboTypes.AddItem "MIDI sequences"
cboTypes.AddItem "MP3's"
cboTypes.AddItem "Windows Media Audio"
cboTypes.AddItem "Real Media (requires Real Alternative)"
cboTypes.AddItem "Quicktime movies (requires Quicktime)"
cboTypes.AddItem "MPEG's"
cboTypes.AddItem "Windows Media Video"
cboTypes.AddItem "PCM Wave files"
cboTypes.AddItem "All files"
File1.Pattern = "*.mp3;*.mp2;*.mp1;*.mid;*.rmi;*.wav;*.rm;*.rmvb;*.mov;*.mpg;*.mpg;*.mpe;*.wma;*.wmv"

If oldPath <> "" Then
  Dir1.path = oldPath & "\"
End If

End Sub


