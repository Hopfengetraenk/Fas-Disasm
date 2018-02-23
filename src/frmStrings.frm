VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStrings 
   BorderStyle     =   5  '?nderbares Werkzeugfenster
   Caption         =   "String & Symbols Window"
   ClientHeight    =   3564
   ClientLeft      =   12048
   ClientTop       =   5052
   ClientWidth     =   3312
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3564
   ScaleWidth      =   3312
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView Lv_Strings 
      Height          =   2052
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "right click (or enter) to  find next"
      Top             =   0
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   3620
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "ID"
         Text            =   "ID"
         Object.Width           =   617
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "mod"
         Text            =   "Modul"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "val"
         Text            =   "Value"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "type"
         Text            =   "Type"
         Object.Width           =   1235
      EndProperty
   End
End
Attribute VB_Name = "frmStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myFormExtenter As FormExt_MostTop
Public LV As Ext_ListView

Private FrmBoarderSize
Private FrmTitleSize


Private Sub Form_Initialize()
   Set myFormExtenter = New FormExt_MostTop
   myFormExtenter.Create Me
   
   Set LV = New Ext_ListView
   LV.Create Lv_Strings
   
   FrmBoarderSize = (getDLGScrollbarSize) * Screen.TwipsPerPixelX
   FrmTitleSize = getDLGIconSize * Screen.TwipsPerPixelX
   
End Sub



Public Sub Clear()
'   Lb_Strings__.Clear
   
   Lv_Strings.ListItems.Clear
End Sub


Private Sub Form_Resize()
   On Error Resume Next
   Lv_Strings.Width = Me.Width - FrmBoarderSize  '220
   Lv_Strings.Height = Me.Height - FrmBoarderSize - FrmTitleSize '500
End Sub



Private Sub FindFirst()

   'Search from Start
   With FrmMain.LV_Log
      Set .SelectedItem = .ListItems(1)

     ' set to next item to continue search
'       Set .SelectedItem = .ListItems(1 + .SelectedItem.Index)
   End With
   
   Dim FindThis$
   FindThis = getSelectedText()
   '   Split(Lb_Strings__, "  ")(2)
   

   FindNext FindThis
End Sub

Public Sub FindNext(FindThis$, Optional bOnEndWarpToStart = True)
   ' "off:" & FasCmdlineObj.Position
'   On Error Resume Next
'   Dim i
   'On Error Resume Next
   
      
   Dim bWasFound As Boolean
   
   With FrmMain.LV_Log
   
      Dim bIsAtTheEnd As Boolean
      bIsAtTheEnd = (.ListItems.count = .SelectedItem.Index)
      
      If bIsAtTheEnd And bOnEndWarpToStart Then
         .ListItems(1).Selected = True
      End If
      

        FrmMain.Panel_Status = "@" & .SelectedItem.Text & " find '" & FindThis & "'"
   

       .ListItems(.ListItems.count).EnsureVisible
      
      
       Dim li As MSComctlLib.ListItem
       Dim si As MSComctlLib.ListSubItem
      
       For Each li In .ListItems
         Dim tmp
         tmp = li.Index + .SelectedItem.Index
         Set li = .ListItems(tmp - (.ListItems.count And (tmp > .ListItems.count)))
         
       ' find in subitems
         For Each si In li.ListSubItems
              bWasFound = InStr(1, si, FindThis, _
                       IIf(FrmSearch.Chk_CaseSensitiv = vbChecked, vbBinaryCompare, vbTextCompare))
                       
              If bWasFound Then Exit For
          Next
          
          If bWasFound Then Exit For
       Next
      'if the Loop gets complete => nothing was found
      '... and => li is set to 'nothing'

      If bWasFound Then
         li.Selected = True
         FrmMain.LV_Log_Ext.EnsureVisible li
'         .SelectedItem.EnsureVisible
         
         FrmMain.Panel_Detail = "found @" & li.Text
      Else
         FrmMain.Panel_Detail = "not found."
      End If
      
      
   End With
End Sub

Private Sub Lv_Strings_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   If KeyAscii = vbKeyReturn Then
      FindNext getSelectedText()
      '(Split(Lb_Strings__, "  ")(0))
   End If
KeyAscii = 0
End Sub
Private Sub Lv_Strings_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Error Resume Next
   If Button And vbRightButton Then
      FindNext getSelectedText()
      '(Split(Lb_Strings__, "  ")(1))
   ElseIf Button And vbLeftButton Then
      FindFirst
   End If
End Sub

Private Function getSelectedText()
   getSelectedText = LV.ListSubItem(Lv_Strings.SelectedItem, "val")
End Function



