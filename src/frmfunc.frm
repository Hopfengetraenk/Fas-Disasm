VERSION 5.00
Begin VB.Form frmFunction 
   BorderStyle     =   5  '?nderbares Werkzeugfenster
   Caption         =   "Function Window"
   ClientHeight    =   3528
   ClientLeft      =   15528
   ClientTop       =   5052
   ClientWidth     =   996
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3528
   ScaleWidth      =   996
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox Lb_Functions 
      Height          =   3120
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Double click to jump"
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmFunction"
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
   Lb_Functions.Clear
End Sub

Private Sub Form_Resize()
   Lb_Functions.Width = Me.Width - 100
   Lb_Functions.Height = Me.Height - 300
End Sub


Private Sub Lb_Functions_Click()

   On Error Resume Next
   
   Dim tmp
   tmp = Split(Lb_Functions, "  ")(1)
   
   FrmMain.LV_Log_Ext.EnsureVisible _
   FrmMain.LV_Log_Ext.OffsetKeyGet(1, tmp)

'   With FrmMain.LV_Log.ListItems
'
'    'Find myItem
'      Dim myItem As ListItem
'      Set myItem = .item(tmp)
'
'    ' Ensure that List Item is always shown on top
'      .item(.count).EnsureVisible
'
'    ' Select it
'      myItem.Selected = True
'
'    ' scroll 6 item up in list (if it's item 1..6)
'      .item(IIf(myItem.Index <= 6, _
'               myItem.Index, _
'               myItem.Index - 6) _
'            ).EnsureVisible
'
'   End With

End Sub




