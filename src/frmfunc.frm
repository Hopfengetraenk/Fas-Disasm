VERSION 5.00
Begin VB.Form frmFunction 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Function Window"
   ClientHeight    =   3510
   ClientLeft      =   11805
   ClientTop       =   390
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox Lb_Functions 
      Height          =   3180
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

Private Sub Form_Load()
   MostTop Me.Hwnd
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
   tmp = "off:" & Split(Lb_Functions, "  ")(1)
   
   With FrmMain.ListView1.ListItems
   
    'Find myItem
      Dim myItem As ListItem
      Set myItem = .item(tmp)
      
    ' Ensure that List Item is always shown on top
      .item(.count).EnsureVisible
      
    ' Select it
      myItem.Selected = True
      
    ' scroll 6 item up in list (if it's item 1..6)
      .item(IIf(myItem.Index <= 6, _
               myItem.Index, _
               myItem.Index - 6) _
            ).EnsureVisible
    
   End With

End Sub




