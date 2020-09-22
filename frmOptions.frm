VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FireAMP Preferences"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8460
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstSkinPaths 
      Height          =   285
      Left            =   360
      TabIndex        =   28
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7080
      TabIndex        =   9
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   6360
      Width           =   975
   End
   Begin MSComctlLib.TreeView trvOptions 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   8916
      _Version        =   393217
      Indentation     =   18
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraUpdate 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2760
      TabIndex        =   38
      Top             =   960
      Width           =   5535
      Begin VB.PictureBox picHolder1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   120
         ScaleHeight     =   4335
         ScaleWidth      =   5295
         TabIndex        =   41
         Top             =   720
         Width           =   5295
         Begin VB.PictureBox picUpdateHolder 
            BorderStyle     =   0  'None
            Height          =   3135
            Left            =   120
            ScaleHeight     =   3135
            ScaleWidth      =   5055
            TabIndex        =   44
            Top             =   1200
            Width           =   5055
            Begin VB.TextBox txtStatus 
               Height          =   1935
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   48
               Top             =   120
               Width           =   4815
            End
            Begin VB.CommandButton cmdDownloadUpdate 
               Caption         =   "&Download Update"
               Enabled         =   0   'False
               Height          =   375
               Left            =   2640
               TabIndex        =   46
               Top             =   2640
               Width           =   2175
            End
            Begin ComctlLib.ProgressBar pbStatus 
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   2160
               Width           =   4695
               _ExtentX        =   8281
               _ExtentY        =   450
               _Version        =   327682
               Appearance      =   0
            End
         End
         Begin VB.Label lblStatus 
            Caption         =   "Checking for newer versions ..."
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   840
            Width           =   4935
         End
         Begin VB.Label lblNewVersion 
            AutoSize        =   -1  'True
            Caption         =   "Available version:"
            Height          =   225
            Left            =   120
            TabIndex        =   43
            Top             =   480
            Width           =   1425
         End
         Begin VB.Label lblCurrentVersion 
            AutoSize        =   -1  'True
            Caption         =   "Current Version:"
            Height          =   225
            Left            =   120
            TabIndex        =   42
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   4560
         TabIndex        =   39
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label11 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame fraVisOptions 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2760
      TabIndex        =   33
      Top             =   960
      Width           =   5535
      Begin ComctlLib.Slider sldCPU 
         Height          =   615
         Left            =   240
         TabIndex        =   37
         Top             =   1440
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1085
         _Version        =   327682
         Max             =   100
         TickStyle       =   2
         TickFrequency   =   10
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Visualization Speed"
         Height          =   225
         Left            =   360
         TabIndex        =   36
         Top             =   1080
         Width           =   1620
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visualizations"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   3960
         TabIndex        =   34
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label Label25 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame fraGeneral 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2760
      TabIndex        =   1
      Top             =   960
      Width           =   5535
      Begin VB.TextBox txtSepChar 
         Height          =   375
         Left            =   4320
         MaxLength       =   1
         TabIndex        =   30
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox chkID3 
         Alignment       =   1  'Right Justify
         Caption         =   "Never warn about missing ID3 Data"
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   960
         Width           =   4815
      End
      Begin VB.Label Label1 
         Caption         =   "Artist/Title Separator:"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   4560
         TabIndex        =   17
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label15 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame fraKeyShort 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   5535
      Begin VB.ListBox lstShortCuts 
         BackColor       =   &H00FFFFFF&
         Height          =   4110
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keyboard Shortcuts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   3465
         TabIndex        =   11
         Top             =   300
         Width           =   1860
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame fraSkins 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   5535
      Begin VB.PictureBox picHolder 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1680
         ScaleHeight     =   495
         ScaleWidth      =   1815
         TabIndex        =   31
         Top             =   960
         Width           =   1815
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "&Browse ..."
            Height          =   375
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1335
         Left            =   240
         TabIndex        =   22
         Top             =   3720
         Width           =   5055
         Begin VB.Label lblComment 
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   4815
         End
         Begin VB.Label lblAuthor 
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   4815
         End
         Begin VB.Label lblName 
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.ListBox lstSkins 
         Height          =   1410
         Left            =   240
         TabIndex        =   21
         Top             =   2160
         Width           =   5055
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Installed Skins:"
         Height          =   225
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   1260
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Load new Skin :"
         Height          =   225
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skins"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   4800
         TabIndex        =   13
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame fraStartUP 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   5535
      Begin VB.CheckBox chkDefSkin 
         Alignment       =   1  'Right Justify
         Caption         =   "Load Default Skin"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1680
         Width           =   4815
      End
      Begin VB.CheckBox chkSplash 
         Alignment       =   1  'Right Justify
         Caption         =   "Show Splash Screen"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   4560
         TabIndex        =   15
         Top             =   300
         Width           =   765
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5295
      End
   End
   Begin MSWinsockLib.Winsock wsUpdate 
      Left            =   3240
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   170
      Top             =   50
      Width           =   750
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000013&
      X1              =   0
      X2              =   8520
      Y1              =   6260
      Y2              =   6260
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   8520
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FireAMP Preferences"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   18
      Top             =   120
      Width           =   3405
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000004&
      X1              =   0
      X2              =   8520
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   0
      X2              =   8520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'credits: my friend

'
'Created By:The Punisher
'Ideasoft, Inc.
'

Option Explicit

Dim spl, addr, hspl, revspl, fn
Dim rhost, rport, rpath, fname, fsize, cindex, reqsize As Boolean

'Private Const ServerAddr As String = "fireamp.phpnet.us/update.txt"
Private ServerAddr As String
Private gotUpdateFile As Boolean

Private newFileName As String, dlComplete As Boolean

Private Sub chkID3_Click()
SaveSetting App.EXEName, "Settings", "Alerts", CStr(Abs(chkID3.Value))
End Sub

Private Sub cmdApply_Click()
Dim Options As FireAMPoptions

Options.loadDefaultSkin = chkDefSkin.Value
Options.showSplashScreen = chkSplash.Value
If txtSepChar.Text = "" Then txtSepChar.Text = "/"
Options.defaultSepChar = CByte(Asc(txtSepChar.Text))

sepChar = Asc(txtSepChar.Text)
If Not Fsys.FileExists(App.path & "\FireAMP.Options") Then Fsys.CreateTextFile App.path & "\FireAMP.Options"
Open App.path & "\FireAMP.Options" For Binary Access Read Write Lock Write As #1
Put #1, , Options
Close #1
End Sub

Private Sub cmdBrowse_Click()
Dim Fout As textStream, i As Integer, d As New Dictionary
frmDummy.cd1.FileName = ""
frmDummy.cd1.Filter = "FireAMP! Skins (*.cfs)|*.cfs"
frmDummy.cd1.ShowOpen

For i = 0 To lstSkins.ListCount - 1
d.Add lstSkinPaths.List(i), lstSkins.List(i)
Next

If frmDummy.cd1.FileName <> "" Then
If Not d.Exists(frmDummy.cd1.FileName) Then
 lstSkins.AddItem frmDummy.cd1.FileTitle
 lstSkinPaths.AddItem frmDummy.cd1.FileName
End If
End If


Set Fout = Fsys.OpenTextFile(App.path & "\FireAMP.Skins.lst", ForWriting, True)

For i = 0 To lstSkins.ListCount - 1
 Fout.WriteLine lstSkins.List(i)
 Fout.WriteLine lstSkinPaths.List(i)
  
Next
Fout.Close
Set Fout = Nothing
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDownloadUpdate_Click()
newFileName = App.path & "\FireAMP.Update.zip"
ServerAddr = "fireamp.phpnet.us/update.zip"
startUpdate
End Sub

Private Sub cmdOK_Click()
cmdApply_Click
Unload Me
End Sub

Private Sub Form_Load()
Dim nodOpt As Node
'Main nodes
Set nodOpt = trvOptions.Nodes.Add(, tvwNext, "Pref", "Preferences")
Set nodOpt = trvOptions.Nodes.Add(, tvwNext, "KbShrt", "Keyboard Shortcuts")
'Child nodes
Set nodOpt = trvOptions.Nodes.Add("Pref", tvwChild, "Gen", "General")
Set nodOpt = trvOptions.Nodes.Add("Pref", tvwChild, "StrUP", "Start Up")
Set nodOpt = trvOptions.Nodes.Add("Pref", tvwChild, "Skn", "Skins")
Set nodOpt = trvOptions.Nodes.Add("Pref", tvwChild, "Vis", "Visualizations")
Set nodOpt = trvOptions.Nodes.Add("Pref", tvwChild, "Upd", "Update")
'Hide all frames
HideFrames
fraGeneral.Visible = True


' initialize shortcuts
lstShortCuts.AddItem "Play Clip -> Space Bar"
lstShortCuts.AddItem "Stop Clip -> S"
lstShortCuts.AddItem "Open Clip ->  O"
lstShortCuts.AddItem "Exit -> X"
lstShortCuts.AddItem "Volume up -> H"
lstShortCuts.AddItem "Volume Down -> G"
lstShortCuts.AddItem "Mute -> M"
lstShortCuts.AddItem "Minimize -> N"
lstShortCuts.AddItem "Change Visualization: Next-> B : Previous-> V"
lstShortCuts.AddItem "Full Screen/Normal video -> F"


If Fsys.FileExists(App.path & "\FireAMP.Options") Then
Open App.path & "\FireAMP.Options" For Binary Access Read As 1
Get #1, , theOptions
Close #1
End If

  chkDefSkin.Value = theOptions.loadDefaultSkin
 chkSplash.Value = theOptions.showSplashScreen
txtSepChar.Text = Chr(CLng(theOptions.defaultSepChar))

chkID3.Value = Abs(Val(GetSetting(App.EXEName, "Settings", "Alerts", "1")))
Dim Fin As textStream
Set Fin = Fsys.OpenTextFile(App.path & "\FireAMP.Skins.Lst", ForReading, True)
While Not Fin.AtEndOfStream
lstSkins.AddItem Fin.readLine
lstSkinPaths.AddItem Fin.readLine
Wend
Fin.Close
Set Fin = Nothing

sldCPU.Value = 100 - sleepFactor
End Sub

Private Sub lstSkins_Click()
Frame2.Visible = True
Dim sn As String, sa As String, snn As String
readSkinData lstSkinPaths.List(lstSkins.ListIndex), sn, sa, snn

lblName.Caption = sn
lblAuthor.Caption = "Created by: " & sa
lblComment.Caption = snn
renderSkin lstSkinPaths.List(lstSkins.ListIndex)
End Sub

Private Sub sldCPU_Scroll()
sleepFactor = 100 - sldCPU.Value

End Sub

Private Sub trvOptions_NodeClick(ByVal Node As MSComctlLib.Node)
'Note you can call "HideFrames" in each case also
HideFrames

Select Case Node.Key
Case "Gen", "Pref"
    fraGeneral.Visible = True
Case "StrUP"
    fraStartUP.Visible = True
Case "Skn"
    fraSkins.Visible = True
Case "KbShrt"
    fraKeyShort.Visible = True
Case "Vis"
    fraVisOptions.Visible = True
Case "Upd"
    fraUpdate.Visible = True
    newFileName = "update.txt"
    ServerAddr = "fireamp.phpnet.us/update.txt"
    startUpdate
Case Default

End Select
End Sub

Private Sub HideFrames()
Dim Control As Object
    For Each Control In Me
        If TypeOf Control Is Frame Then
            Control.Visible = False
        End If
    Next
End Sub

Sub startUpdate()
If Fsys.FileExists(App.path & "\" & newFileName) Then Fsys.DeleteFile App.path & "\" & newFileName
lblCurrentVersion.Caption = "Current Version: " & App.Major & "." & App.Minor & "." & App.Revision

       revspl = Split(StrReverse(ServerAddr), "/")
        fname = StrReverse(revspl(0))
        
        spl = Split(ServerAddr, "/")
        hspl = Split(spl(0) & ":80", ":")
        rhost = hspl(0)
        rport = hspl(1)
        rpath = Mid(ServerAddr, Len(spl(0)) + 1)

txtStatus.Text = ""

logText "Looking up Update.txt ..."
logText "Connecting to " & rhost & ":" & rport & "..."


wsUpdate.Close
wsUpdate.Connect rhost, rport
reqsize = True
    cindex = 0
End Sub

Private Sub wsUpdate_Connect()
    Dim request
    If reqsize Then
        request = Replace(Replace(HTTPGET, "{HOST}", rhost), "{PATH}", "http://" & ServerAddr)
        logText "Connected to " & wsUpdate.RemoteHost & ":" & wsUpdate.RemotePort
        logText "Requesting file size..."
        logText "request:" & vbCrLf & request
        wsUpdate.SendData request
        logText "Size request sent. Awaiting response..."
    Else
        logText "Requesting file from position " & cindex - fsize & "..."
        request = Replace(Replace(Replace(HTTPGETRANGE, "{HOST}", rhost), "{PATH}", "http://" & ServerAddr), "{RANGE}", cindex - fsize)
        logText "request:" & vbCrLf & request
        wsUpdate.SendData request
        logText "Downloading..."
    End If
    
End Sub

Private Sub wsUpdate_DataArrival(ByVal bytesTotal As Long)
    Dim dat As String
    Dim crspl, i, lenspl, dblcrspl, flen, cmpname, cmpext, cmpspl, pcnt
    
    '// If reqsize is true, then parse the HTTP response for the Content-Length value
    If reqsize = True Then
        wsUpdate.GetData dat, , 1024
        
        '// HTTP 404 recieved from the server; abort.
        If LCase(Mid(dat, 1, 12)) = "http/1.1 404" Then
            logText "Recieved HTTP not found. Aborting."
            wsUpdate.Close
            
            Exit Sub
        End If
        
        crspl = Split(dat & vbCrLf & "[EOF]", vbCrLf)
        i = 0
        While crspl(i) <> "[EOF]"
            If LCase(Mid(crspl(i) & "XXXXXXXXXXXXXXX", 1, 15)) = "content-length:" Then
                lenspl = Split(crspl(i) & ": 0", ": ")
                fsize = lenspl(1)
                logText "File size is: " & (fsize \ 1024) & " Kb"
                reqsize = False
            End If
            i = i + 1
        Wend
        
        '// Failed to parse Content-Length value. Invalid data. Abort.
        If reqsize = True Then
            logText "Recieved unsupported or non-HTTP data from the server. Aborting."
            wsUpdate.Close
        Exit Sub
        End If
        
        wsUpdate.Close
        logText "Reconnecting to server..."
        wsUpdate.Connect rhost, rport
        Exit Sub
    Else
    
        '// Otherwise, if reqsize is not true, then append the incomming data to the currently downloading file.
        wsUpdate.GetData dat
    End If
    
    '// Remove the HTTP header
    If LCase(Mid(dat, 1, 12)) = "http/1.1 206" Then
        dblcrspl = Split(dat, vbCrLf & vbCrLf, 2)
        dat = dblcrspl(1)
    End If
    
    '// Append the currently downloading file
    Open App.path & "\" & newFileName For Binary Access Write As #1
        Put #1, LOF(1) + 1, dat 'Mid(dat, 1, Len(dat) - 2)
    Close #1
    
    '// Check the file size against the servers file size to determine wether it has completed downloading
    flen = FileLen(App.path & "\" & newFileName)
    If Val(flen) = Val(fsize) Then
        If Not gotUpdateFile Then
        gotUpdateFile = True
            checkIfNewVersionAvailable
        End If
        '// Download has completed. Notify user, rename .downloading file, etc
        If Not dlComplete Then logText "Update completed."
        
        '// If the .downloading file cannot be renamed because a file of that name already exists,
        '// then append a number to the file name. Eg. foobar(1).exe. (Increase the number until
        '// the file name is unique)
        wsUpdate.Close
        Exit Sub
    End If
    
    '// Calculate & display progress in %
    On Error Resume Next
    pcnt = flen / fsize * 100
    pbStatus.Value = Math.Round(pcnt)
End Sub

Private Sub wsUpdate_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

logText "ERROR: " & Description
wsUpdate.Close
End Sub

Sub logText(Text As String)
txtStatus.Text = txtStatus.Text & Text & vbNewLine
txtStatus.SelStart = Len(txtStatus.Text)
txtStatus.SelLength = 1
End Sub

Public Sub checkIfNewVersionAvailable()

Dim newVersion As String, Fin As textStream
Set Fin = Fsys.OpenTextFile(App.path & "\Update.txt", ForReading, False)
newVersion = Fin.ReadAll

Dim parts() As String
parts = Split(newVersion, ".")

If App.Revision < Val(parts(2)) Then
 logText "A new version is available for download. Click the Download button to install it."
 cmdDownloadUpdate.Enabled = True
Else
 logText "No new versions at this moment. Check again in a few days."
End If
Fin.Close
Set Fin = Nothing
End Sub
