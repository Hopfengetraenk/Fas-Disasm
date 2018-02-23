VERSION 5.00
Begin VB.Form frmlog 
   BorderStyle     =   5  '?nderbares Werkzeugfenster
   Caption         =   "Log Window"
   ClientHeight    =   684
   ClientLeft      =   132
   ClientTop       =   9132
   ClientWidth     =   11808
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   684
   ScaleWidth      =   11808
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox listLog 
      Height          =   624
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Righclick  to clear."
      Top             =   0
      Width           =   11892
   End
End
Attribute VB_Name = "frmlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myFormExtenter As FormExt_MostTop
Private Sub Form_Initialize()
   Set myFormExtenter = New FormExt_MostTop
   myFormExtenter.Create Me
End Sub

Public Sub Clear()
   listLog.Clear
End Sub

Private Sub Form_Resize()
   listLog.Width = Me.Width - 100
   listLog.Height = Me.Height - 300
End Sub


Private Sub listLog_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
      If Button = vbRightButton Then
      Me.Clear
   ElseIf Button = vbRightButton Then
   Else
   End If
   
End Sub


