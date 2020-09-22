VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmCompiler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FireAMP! Skin Complier"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   Icon            =   "Compiler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   8175
      Begin RichTextLib.RichTextBox txtStatus 
         Height          =   4455
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   7858
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"Compiler.frx":11C2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "&Compile"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   7680
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Files:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   8055
      Begin VB.TextBox txtSkin 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   840
         Width           =   5415
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse ..."
         Height          =   375
         Left            =   6600
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtSrc 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Output File:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Script File:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8400
      Y1              =   1096
      Y2              =   1096
   End
   Begin VB.Line Line1 
      X1              =   3451
      X2              =   3451
      Y1              =   0
      Y2              =   1095
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   0
      Picture         =   "Compiler.frx":123D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "FireAMP! Skin Compiler"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   3555
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "frmCompiler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isError As Boolean
Dim Fsys As New FileSystemObject
Private Sub checkScript()

Dim FIn As TextStream, skinNames(0 To 8) As String
Dim j As Integer, t As Long
t = Timer
txtStatus.Text = ""
Set FIn = Fsys.OpenTextFile(txtSrc.Text)

txtStatus.SelColor = &H800080
writeStatus "Parsing script ..."
txtStatus.SelColor = 10485760
writeStatus "Parsing Region 'files' ..."

Dim line$, ln As Integer
ln = 1
Dim parts() As String
While Not FIn.AtEndOfStream
line = FIn.ReadLine
ln = ln + 1

' check for files
If LCase(line) = "files" Then
FIn.ReadLine

While line <> "%>"

line = FIn.ReadLine
  ln = ln + 1
  parts = Split(line, ":")
  If UBound(parts) > 0 Then
     If Dir(Trim(parts(1))) = "" Then
     
       txtStatus.SelColor = 128
       writeStatus "File not found: " & parts(1), ln
     Else
       txtStatus.SelColor = 8421440
       writeStatus "  [Resource file: " & parts(1) & "]"
       skinNames(j) = Mid(txtSrc.Text, 1, Len(txtSrc.Text) - Len(Fsys.GetFileName(txtSrc.Text))) & Trim(parts(1))
       
       j = j + 1
     End If
  End If

 Wend

End If

' check data
If LCase(line) = "data" Then
txtStatus.SelColor = 10485760
writeStatus "Parsing Region 'data'..."

FIn.ReadLine
While line <> "%>"
 line = FIn.ReadLine
 ln = ln + 1
parts = Split(line, ":")
  If UBound(parts) > 0 Then
     If Not checkData(parts(0)) Then
       txtStatus.SelColor = 128
       writeStatus "  Syntax error: " & parts(0), ln
       End If
  End If
Wend

End If

' check arrange
If LCase(line) = "arrange" Then
txtStatus.SelColor = 10485760
writeStatus "Parsing Region 'arrange' ..."

FIn.ReadLine
While line <> "%>"
 line = FIn.ReadLine
 ln = ln + 1
parts = Split(line, ":")
  If UBound(parts) > 0 Then
     If Not checkArrange(parts(0)) Then
       txtStatus.SelColor = 128
       writeStatus "  Syntax error: " & parts(0), ln
       End If
  End If
Wend
End If
' check fonts
If LCase(line) = "fonts" Then
txtStatus.SelColor = 10485760
writeStatus "Parsing Region 'fonts' ..."

FIn.ReadLine
While line <> "%>"
 line = FIn.ReadLine
 ln = ln + 1
parts = Split(line, ":")
  If UBound(parts) > 0 Then
     If Not checkFonts(parts(0)) Then
       txtStatus.SelColor = 128
       writeStatus "  Syntax error: " & parts(0), ln
       End If
  End If
Wend
End If

' check colors
If LCase(line) = "colors" Then
txtStatus.SelColor = 10485760
writeStatus "Parsing Region 'colors' ..."

FIn.ReadLine
While line <> "%>"
 line = FIn.ReadLine
 ln = ln + 1
parts = Split(line, ":")
  If UBound(parts) > 0 Then
     If Not checkColors(parts(0)) Then
       txtStatus.SelColor = 128
       writeStatus "  Syntax error: " & parts(0), ln
       End If
  End If
Wend
End If
Wend
writeStatus ""
txtStatus.SelColor = &H800080

If Not isError Then
writeStatus "Parse Succeeded."
writeStatus ""
txtStatus.SelColor = &H800080
writeStatus "Building Skin ..."

skinNames(8) = txtSrc.Text
' the skinner
Dim skin As New FireSkinLibrary.FireSkinner
skin.makeFireSkin skinNames, txtSkin.Text
Set skin = Nothing

txtStatus.SelColor = &H800080
writeStatus "Build Succeeded: " & Round(Abs(Timer - t), 2) & "s, " & ln & " lines"
txtStatus.SelColor = &H800080
writeStatus "Skin size: " & CInt(FileLen(txtSkin.Text) / 1024&) & "kb" & " (" & Round(FileLen(txtSkin.Text) / (1024& * 1024&), 2) & "MB)"
MsgBox "Build Succeeded!", vbOKOnly + vbInformation, "Finished building"

Else
writeStatus "Parse failed. There are errors"
txtStatus.SelColor = vbRed
writeStatus "Cannot build skin."
Beep
End If


End Sub
Private Sub cmdBrowse_Click()
cd1.CancelError = False
cd1.Filter = "FireSkin Specification(*.fss)|*.fss"
cd1.ShowOpen
If cd1.FileName <> "" Then
txtSrc.Text = cd1.FileName
txtSkin.Text = Mid(txtSrc.Text, 1, Len(txtSrc.Text) - Len(Fsys.GetExtensionName(txtSrc.Text))) & "cfs"
End If
End Sub

Private Sub cmdCompile_Click()
If Trim(txtSrc.Text) <> "" Then
isError = False
checkScript
End If
End Sub

Private Sub cmdExit_Click()
End
End Sub

' check data section
Function checkData(ByVal line$) As Boolean
checkData = True
line = LCase(line)
If line Like "name" Then
ElseIf line Like "author" Then
ElseIf line Like "notes" Then
ElseIf line Like "%>" Then
Else
checkData = False
End If

End Function

Sub writeStatus(StatusStr As String, Optional lineNo As Integer)
Dim status As String
status = StatusStr
If lineNo <> 0 Then
status = status & " at line#" & Format(lineNo, "000")
isError = True
End If
txtStatus.SelText = txtStatus.SelText & status & vbNewLine
End Sub

Function checkArrange(line$) As Boolean
checkArrange = True
line = LCase(line)
If line Like "main-caption" Then
ElseIf line Like "main-seek-bar" Then
ElseIf line Like "main-play-button" Then
ElseIf line Like "main-stop-button" Then
ElseIf line Like "main-open-button" Then
ElseIf line Like "main-close-button" Then
ElseIf line Like "main-min-button" Then
ElseIf line Like "main-time" Then
ElseIf line Like "main-info" Then
ElseIf line Like "pl-caption" Then
ElseIf line Like "pl-list" Then
ElseIf line Like "pl-bar" Then
ElseIf line Like "song-title" Then
ElseIf line Like "song-album" Then
ElseIf line Like "vis" Then
ElseIf line Like "vis-box" Then
ElseIf line Like "video" Then
Else
checkArrange = False
End If
End Function

Function checkFonts(line$) As Boolean
checkFonts = True
If line Like "main-caption" Then
ElseIf line Like "main-title" Then
ElseIf line Like "main-album" Then
ElseIf line Like "main-time" Then
ElseIf line Like "main-info" Then
ElseIf line Like "pl-caption" Then
ElseIf line Like "pl-list" Then
Else
checkFonts = False
End If
End Function

Function checkColors(line$) As Boolean
checkColors = True
If line Like "main-caption" Then
ElseIf line Like "main-title" Then
ElseIf line Like "main-album" Then
ElseIf line Like "main-time" Then
ElseIf line Like "main-info" Then
ElseIf line Like "main-album-back" Then
ElseIf line Like "main-title-back" Then
ElseIf line Like "pl-caption" Then
ElseIf line Like "pl-list" Then
ElseIf line Like "pl-back-list" Then
ElseIf line Like "pl-current" Then
Else
checkColors = False
End If
End Function

Private Sub Form_Unload(Cancel As Integer)
Set Fsys = Nothing
End Sub
