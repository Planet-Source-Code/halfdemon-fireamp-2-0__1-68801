VERSION 5.00
Begin VB.Form frmExtVis 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Visualizer"
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   -30
   ClientWidth     =   3585
   Icon            =   "frmExtVis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   239
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmExtVis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Unload Me
End Sub

'
' external visualizations
'

' last updated: 2006 May 05, Humanoid

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
 PopupMenu frmDummy.mnuVis
End If

End Sub

