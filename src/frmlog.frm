VERSION 5.00
Begin VB.Form frmlog 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Log Window"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   10455
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox listLog 
      Height          =   3180
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Righclick  to clear."
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub clear()
   listLog.clear
End Sub

Private Sub Form_Resize()
   listLog.Width = Me.Width - 100
   listLog.Height = Me.Height - 300
End Sub


Private Sub listLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
      If Button = vbRightButton Then
      Me.clear
   ElseIf Button = vbRightButton Then
   Else
   End If
   
End Sub


