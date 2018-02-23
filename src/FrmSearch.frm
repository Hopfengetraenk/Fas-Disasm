VERSION 5.00
Begin VB.Form FrmSearch 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Search"
   ClientHeight    =   780
   ClientLeft      =   8316
   ClientTop       =   1056
   ClientWidth     =   3192
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   3192
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Chk_CaseSensitiv 
      Caption         =   "Case Sensitiv"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox Chk_AutoSearch 
      Caption         =   "Autosearch"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Value           =   1  'Aktiviert
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2985
   End
End
Attribute VB_Name = "FrmSearch"
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

Private Sub Combo1_Change()

If Chk_AutoSearch = vbUnchecked Then Exit Sub

On Error GoTo Combo1_Change_err
      Me.Caption = Me.Tag & " - in Process"
      frmStrings.FindNext Combo1.Text
      Me.Caption = Me.Tag
Exit Sub
Combo1_Change_err:
   Me.Caption = Me.Tag & " - String not found"
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   
   On Error GoTo Cmd_Search_Click_err
   With Combo1
      frmStrings.FindNext .Text
       
      Dim item
      
      For item = 0 To .ListCount
         If .List(item) = .Text Then
            Exit Sub
         End If
      Next
      
      .AddItem .Text
   End With
   Me.Caption = Me.Tag
End If
Exit Sub
Cmd_Search_Click_err:
   Me.Caption = Me.Tag & " - String not found"
End Sub

Private Sub Form_Load()
   Me.Tag = Me.Caption
   
 '  MostTop Me.Hwnd
End Sub
