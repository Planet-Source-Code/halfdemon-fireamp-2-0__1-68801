VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDummy 
   Caption         =   "frmDummy"
   ClientHeight    =   135
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   9
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5760
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFireAMP 
      Caption         =   "FireAMP"
      WindowList      =   -1  'True
      Begin VB.Menu mnuAbout 
         Caption         =   "About FireAMP!"
      End
      Begin VB.Menu mnuNull0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences ..."
      End
      Begin VB.Menu mnuShowPlaylist 
         Caption         =   "Show Playlist"
      End
      Begin VB.Menu mnuNull1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPlaylist 
      Caption         =   "Playlist"
      Begin VB.Menu mnuTagEdit 
         Caption         =   "Edit Tag ..."
      End
      Begin VB.Menu mnuViewInfo 
         Caption         =   "Clip Properties ..."
      End
      Begin VB.Menu mnuNull2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenPL 
         Caption         =   "Open Playlist ..."
      End
      Begin VB.Menu mnuSavePL 
         Caption         =   "Save Playlist ..."
      End
      Begin VB.Menu mnuNull4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoogleSearch 
         Caption         =   "Google Search"
         Begin VB.Menu mnuSearchArtist 
            Caption         =   "Search for Artist"
         End
         Begin VB.Menu mnuSearchForAlbum 
            Caption         =   "Search for Album"
         End
      End
      Begin VB.Menu mnuPlaylistOptions 
         Caption         =   "Playlist"
         Begin VB.Menu mnuRepeatPlaylist 
            Caption         =   "Repeat Playlist"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuRandomPlayback 
            Caption         =   "Random Playback"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuNull6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuClearPlaylist 
            Caption         =   "Clear Playlist"
         End
         Begin VB.Menu mnuSearch 
            Caption         =   "Find Clip in Playlist ..."
         End
         Begin VB.Menu mnuAddClips 
            Caption         =   "Add Clips ..."
         End
      End
      Begin VB.Menu mnuSepClose 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClosePlaylist 
         Caption         =   "Hide Playlist"
      End
   End
   Begin VB.Menu mnuVis 
      Caption         =   "Visualization"
      Begin VB.Menu mnuGroupName 
         Caption         =   "Visualizations"
         Begin VB.Menu mnuObjectName 
            Caption         =   ""
            Checked         =   -1  'True
            Index           =   0
         End
      End
      Begin VB.Menu mnunull3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadPi 
         Caption         =   "Import Visualization list ..."
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "Configure this visualization ..."
      End
      Begin VB.Menu mnuNull5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdjustLevels 
         Caption         =   "Adjust Levels*"
      End
   End
End
Attribute VB_Name = "frmDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' dummy form to hold menus, this form is never shown
'

' last updated: 2006 May 05, Humanoid

' show about form
Private Sub mnuAbout_Click()

frmAbout.Show vbModal
End Sub

' show the add clips dialog
Private Sub mnuAddClips_Click()
frmAddClips.Show vbModal
End Sub

' show sndvol32.exe for stereo mix
Public Sub mnuAdjustLevels_Click()
Dim a As Double
MsgBox "Select Stereo Mix if necessary" & vbNewLine & "Adjust the Stereo Mix Level to just above zero" & vbNewLine & "If Stereo Mix does not appear, re-run this option or select some other device (like mic)", vbInformation + vbOKOnly, "Adjust levels"
 a = Shell("sndvol32.exe", vbNormalFocus)
SendKeys "%p"
SendKeys "{ENTER}"
SendKeys "%r"
SendKeys "{TAB}"
SendKeys "{DOWN}"
SendKeys " "
SendKeys "{ENTER}"

End Sub

' dialog to change the skin
Public Sub mnuChangeSkin_Click()
cd1.FileName = ""
cd1.Filter = "FireAMP! Skins (*.cfs)|*.cfs"
cd1.ShowOpen
If cd1.FileName <> "" Then
 renderSkin cd1.FileName
End If

End Sub

' clear playlist
Private Sub mnuClearPlaylist_Click()
 frmFirePL.lstPaths.Clear
 frmFirePL.lstPl.ListItems.Clear
 'frmFireMain.lstPl.ListItems.Clear
 
 frmFirePL.lblCaption.Caption = "PlayList"
End Sub

Private Sub mnuClosePlaylist_Click()
frmFirePL.Hide
'frmFireMain.picPlaylist.Visible = False
End Sub

' configure dialog for visualization
Private Sub mnuConfig_Click()
On Error Resume Next ' just in case ...
oPlugIn.doConfig
End Sub

' exit
Private Sub mnuExit_Click()

End

End Sub

' import visualization list
Public Sub mnuLoadPi_Click()

' show the file chooser
cd1.CancelError = False
cd1.Filter = "Orange Soda Visualization List|*.soda"
cd1.FileName = ""
cd1.ShowOpen

Dim t As textStream, d As New Dictionary, s As String
Set t = Fsys.OpenTextFile(App.path & "\Sodas.list", ForReading, True)
While Not t.AtEndOfStream
s = t.readLine
d.Add s, s
Wend

t.Close
Set t = Fsys.OpenTextFile(App.path & "\Sodas.list", ForAppending, False)
If Not d.Exists(cd1.FileName) Then
t.WriteLine cd1.FileName
modPi.importXPIList cd1.FileName
End If
t.Close
Set t = Nothing

End Sub

Private Sub mnuMakeHTMLPlaylist_Click()
modPlaylist.toHTMLPlaylist
End Sub

Private Sub mnuObjectName_Click(Index As Integer)
'On Error Resume Next
frmFireMain.tmrVisUpdate.Enabled = False
DoStop
 'Set oPlugIn = Nothing
 'Set oPlugIn = CreateObject(Split(mnuObjectName(Index).tag, ",")(0))
 mnuConfig.Enabled = CBool(Split(mnuObjectName(Index).tag, ",")(1))
 Dim i As Integer
 For i = 0 To mnuObjectName.Count - 1
 mnuObjectName(i).Checked = False
 Next i
 mnuObjectName(Index).Checked = True
 SaveSetting App.EXEName, "Visualization", "Object", Split(mnuObjectName(Index).tag, ",")(0)
 SaveSetting App.EXEName, "Visualization", "Config", mnuConfig.Enabled
 SaveSetting App.EXEName, "Visualization", "Name", mnuObjectName(Index).Caption
 frmFireMain.setupVisualization
 frmFireMain.tmrVisUpdate.Enabled = True
End Sub

' open media
Private Sub mnuOpenMedia_Click()
frmFireMain.picBtn_Click 2
End Sub

' open playlist
Private Sub mnuOpenPL_Click()
cd1.FileName = ""
cd1.Filter = "FireAMP! Playlists(*.fpl)|*.fpl|M3U Playlists|*.m3u|PLS Playlists|*.pls"
cd1.ShowOpen
'frmFireMain.lstPl.ListItems.Clear
If cd1.FileName <> "" Then
If getFileExtensionFromPath(cd1.FileName) = "fpl" Then
openPlayList frmFirePL.lstPl, frmFirePL.lstPaths, cd1.FileName
ElseIf getFileExtensionFromPath(cd1.FileName) = "m3u" Then
 readM3UFile cd1.FileName, frmFirePL.lstPl, frmFirePL.lstPaths
ElseIf getFileExtensionFromPath(cd1.FileName) = "pls" Then
 loadPLS cd1.FileName, frmFirePL.lstPl, frmFirePL.lstPaths
 
End If
frmFirePL.lblCaption.Caption = "Playlist- " & getFileTitleFromPath(cd1.FileName)

End If
'Dim i As Integer
'For i = 1 To frmFirePL.lstPl.ListItems.Count
'frmFireMain.lstPl.ListItems.Add , , frmFirePL.lstPl.ListItems(i).Text
'Next

End Sub

' show the preferences page
Private Sub mnuPreferences_Click()

    frmOptions.Show vbModal

End Sub

Private Sub mnuRandomPlayback_Click()
doRandomPlayback = Not doRandomPlayback
mnuRandomPlayback.Checked = doRandomPlayback
SaveSetting App.EXEName, "Settings", "Playlist Random", CStr(doRepeatPlaylist)
SaveSetting App.EXEName, "Settings", "Playlist Repeat", CStr(False)
mnuRepeatPlaylist.Checked = False
doRepeatPlaylist = False
End Sub

Private Sub mnuRemove_Click()
On Error Resume Next
frmFirePL.lstPaths.RemoveItem frmFirePL.lstPl.SelectedItem.Index - 1
frmFirePL.lstPl.ListItems.Remove frmFirePL.lstPl.SelectedItem.Index
End Sub

' toggle for playlist repeat
Private Sub mnuRepeatPlaylist_Click()
doRepeatPlaylist = Not doRepeatPlaylist
mnuRepeatPlaylist.Checked = doRepeatPlaylist
SaveSetting App.EXEName, "Settings", "Playlist Repeat", CStr(doRepeatPlaylist)
SaveSetting App.EXEName, "Settings", "Playlist Random", CStr(False)
mnuRandomPlayback.Checked = False
doRandomPlayback = False
End Sub

' save the playlist
Private Sub mnuSavePL_Click()
cd1.FileName = ""
cd1.Filter = "FireAMP! Playlists(*.fpl)|*.fpl"
cd1.ShowSave
If cd1.FileName <> "" Then
savePlayList frmFirePL.lstPaths, cd1.FileName
End If

End Sub

' search the playlist
Private Sub mnuSearch_Click()
frmSearch.Show
End Sub

' google search artist
Private Sub mnuSearchArtist_Click()
Dim search As String
search = "www.google.com/search?q=" & Replace(Trim(mnuSearchArtist.Caption), " ", "%20")
ShellExecute frmFireMain.hwnd, "open", search, vbNullString, vbNullString, 5
End Sub

' google search for album
Private Sub mnuSearchForAlbum_Click()
Dim search As String
search = "www.google.com/search?q=" & Replace(Trim(mnuSearchForAlbum.Caption), " ", "%20")
ShellExecute frmFireMain.hwnd, "open", search, vbNullString, vbNullString, 5

End Sub

Private Sub mnuShowPlaylist_Click()
frmFirePL.Show
'frmFireMain.picPlaylist.Visible = True
End Sub

' show the tag editor
Private Sub mnuTagEdit_Click()

If frmFirePL.lstPl.ListItems.Count > 0 Then
frmTagEditor.importCurrentPlaylist
selectedIndex = frmFirePL.lstPl.SelectedItem.Index
frmTagEditor.lstFiles.ListIndex = frmFirePL.lstPl.SelectedItem.Index - 1
frmTagEditor.lstFiles_Click
    frmTagEditor.Visible = True
End If
    
End Sub

Private Sub mnuViewInfo_Click()
' show file information
On Error GoTo JMP
frmProperties.readProperties frmFirePL.lstPaths.List(frmFirePL.lstPl.SelectedItem.Index - 1)
frmProperties.Show
JMP:
End Sub

' show the media libaray
Private Sub mnuViewMediaLib_Click()
DoEvents
frmMediaLib.Show
End Sub
