VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configure LED Bars"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Peak Color"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   3975
      Begin MSComctlLib.Slider sldPR 
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         Max             =   255
         TickStyle       =   3
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider sldPG 
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         Max             =   255
         TickStyle       =   3
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider sldPB 
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         Max             =   255
         TickStyle       =   3
         TickFrequency   =   10
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Red"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Green"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Blue"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   315
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bar Color"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin MSComctlLib.Slider sldBR 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         Max             =   255
         TickStyle       =   3
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider sldBG 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         Max             =   255
         TickStyle       =   3
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider sldBB 
         Height          =   375
         Left            =   840
         TabIndex        =   2
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         Max             =   255
         TickStyle       =   3
         TickFrequency   =   10
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Blue"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Green"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Red"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   300
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

sldPR.Value = PeakRed
sldPG.Value = PeakGreen
sldPB.Value = PeakBlue

sldBB.Value = BarBlue
sldBG.Value = BarGreen
sldBR.Value = BarRed

End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.EXEName, "Settings", "Color1", BarRed
SaveSetting App.EXEName, "Settings", "Color2", BarGreen
SaveSetting App.EXEName, "Settings", "Color3", BarBlue

SaveSetting App.EXEName, "Settings", "Color4", PeakRed
SaveSetting App.EXEName, "Settings", "Color5", PeakGreen
SaveSetting App.EXEName, "Settings", "Color6", PeakBlue
End Sub


Private Sub sldBB_Change()
BarBlue = sldBB.Value
End Sub

Private Sub sldBR_Change()
BarRed = sldBR.Value
End Sub

Private Sub sldBG_Change()
BarGreen = sldBG.Value
End Sub

Private Sub sldPB_Change()
PeakBlue = sldPB.Value
End Sub

Private Sub sldPG_Change()
PeakGreen = sldPG.Value
End Sub

Private Sub sldPR_Change()
PeakRed = sldPR.Value
End Sub
