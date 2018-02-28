VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormExt_MostTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const HWND_NOTOPMOST = -2
Private Windows_ZOrder As Long
Private Const Windows_ZOrder_XORSwitchvalue = HWND_NOTOPMOST Xor HWND_TOPMOST


Private WithEvents mForm  As Form
Attribute mForm.VB_VarHelpID = -1
Public LastState As Boolean


Private mFormCaption$

Private Sub Class_Initialize()
   Windows_ZOrder = HWND_TOPMOST
   
 'State with MostTop enabled
 
 #If DoDebug = 1 Then
   isMostTop = False
 #Else
   isMostTop = True
 #End If
End Sub

Public Sub Create(TargetForm As Form)
   Set mForm = TargetForm
End Sub

Public Property Get isMostTop() As Boolean
  ' Check for 'Windows_ZOrder Switch bits
   isMostTop = Windows_ZOrder And Windows_ZOrder_XORSwitchvalue
End Property

Public Property Let isMostTop(ByVal vNewValue As Boolean)
   If vNewValue Then
      BitHelper_Set Windows_ZOrder, Windows_ZOrder_XORSwitchvalue
   Else
      BitHelper_Clear Windows_ZOrder, Windows_ZOrder_XORSwitchvalue
   End If
   
End Property

Private Sub BitHelper_Set(ByRef IO_Data, Mask)
   IO_Data = IO_Data Or Mask
End Sub

Private Sub BitHelper_Clear(ByRef IO_Data, Mask)
   IO_Data = IO_Data And Not Mask
End Sub



Public Sub Switch_MostTop_NotMost_Window()
   ' Toggle State
    isMostTop = Not isMostTop
   
   
   If Update_MostTop_NotMost_Window = False Then
    'Upps Error 'Toggle back' state
     isMostTop = isMostTop Xor 1

   End If
End Sub


Public Function Update_MostTop_NotMost_Window() As Boolean
   
   
   Dim retval&
   retval = SetWindowPos(mForm.Hwnd, Windows_ZOrder, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
   If retval = 1 Then
       Update_MostTop_NotMost_Window = True

   '   if err.LastDllError
      mForm.Caption = mFormCaption & IIf(isMostTop, _
                           " [Pinned]", "" _
      )
      
   
   End If
   
End Function


Private Sub mForm_Activate()
   If LastState <> isMostTop Then
      Update_MostTop_NotMost_Window
      LastState = isMostTop
   End If
End Sub

Private Sub mForm_Load()
   mFormCaption = mForm.Caption

End Sub
Private Sub mForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Dim Unloadreason As QueryUnloadConstants
   Unloadreason = UnloadMode
   
   Switch_MostTop_NotMost_Window
   
   If Unloadreason = vbFormControlMenu Then Cancel = True
End Sub

