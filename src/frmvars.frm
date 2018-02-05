VERSION 5.00
Begin VB.Form frmStrings 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "String & Symbols Window"
   ClientHeight    =   3120
   ClientLeft      =   11730
   ClientTop       =   5445
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox Lb_Strings 
      Height          =   3180
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "right click (or enter) to  find next"
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const HWND_NOTOPMOST = -2
Private Windows_ZOrder As Long
Private Const Windows_ZOrder_XORSwitchvalue = HWND_NOTOPMOST Xor HWND_TOPMOST

Private Sub Form_Initialize()
   Windows_ZOrder = HWND_TOPMOST
End Sub

Private Sub Form_Load()
   On Error Resume Next

   MostTop Me.Hwnd
  
End Sub


Public Sub Clear()
   Lb_Strings.Clear
End Sub


Private Sub Switch_MostTop_NotMost_Window()
   Windows_ZOrder = Windows_ZOrder Xor Windows_ZOrder_XORSwitchvalue
   SetWindowPos Me.Hwnd, Windows_ZOrder, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim Unloadreason As QueryUnloadConstants
   Unloadreason = UnloadMode
   
   Switch_MostTop_NotMost_Window
   
   If Unloadreason = vbFormControlMenu Then Cancel = True
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Lb_Strings.Width = Me.Width - 100
   Lb_Strings.Height = Me.Height - 300
End Sub



Private Sub Lb_Strings_Click()
   On Error Resume Next
   FindFirst
End Sub
Private Sub FindFirst()

   With FrmMain.ListView1
      Set .SelectedItem = .ListItems(1)
   End With
   FindNext (Split(Lb_Strings, "  ")(0))
End Sub
Public Sub FindNext(Findthis$)
   ' "off:" & FasCmdlineObj.Position
'   On Error Resume Next
'   Dim i
   'On Error Resume Next
   
   With FrmMain.ListView1

       .ListItems(.ListItems.count).EnsureVisible
      
      
       Dim li As MSComctlLib.ListItem
       Dim si As MSComctlLib.ListSubItem
      
       For Each li In .ListItems
         Dim tmp
         tmp = li.Index + .SelectedItem.Index
         Set li = .ListItems(tmp - (.ListItems.count And (tmp > .ListItems.count)))
         
         For Each si In li.ListSubItems
              Dim Cancel As Boolean
              Cancel = InStr(1, si, Findthis, _
                       IIf(FrmSearch.Chk_CaseSensitiv = vbChecked, vbBinaryCompare, vbTextCompare))
              If Cancel Then Exit For
          Next
          If Cancel Then Exit For
       Next

      li.Selected = True
      .SelectedItem.EnsureVisible
   End With
End Sub

Private Sub Lb_Strings_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   If KeyAscii = vbKeyReturn Then
      FindNext (Split(Lb_Strings, "  ")(0))
   End If
KeyAscii = 0
End Sub

Private Sub Lb_Strings_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If Button And vbRightButton Then
      FindNext (Split(Lb_Strings, "  ")(0))
   End If
End Sub
