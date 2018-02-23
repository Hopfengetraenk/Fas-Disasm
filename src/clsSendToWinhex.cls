VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSendToWinhex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private ProcID&                'ProcessID of HWORKS32.exe
Private HWORKS_Path
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbSize As Long) As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

 
Private Type KEYBDINPUT
   wVk As Integer
   wScan As Integer
   dwFlags As Long
   time As Long
   dwExtraInfo As Long
End Type
 
 
Private Type INPUT_TYPE
   dwType As Long
   xi As KEYBDINPUT
End Type
 
 
 
'KEYBDINPUT dwFlags-Konstanten
Private Const KEYEVENTF_EXTENDEDKEY = &H1 'Der Scancode hat das Präfix &HE0
Private Const KEYEVENTF_KEYUP = &H2 'Die angegebene Taste wird losgelassen
Private Const KEYEVENTF_UNICODE = &H4 'Benutzt ein Unicode Buchstaben der nicht von einen der Tastaturcodes stammt welcher eine Tastatureingabe Simuliert
 
'INPUT_TYPE dwType-Konstanten
Private Const INPUT_MOUSE = 0 'Mauseingabe
Private Const INPUT_KEYBOARD = 1 'Tastatureingabe
Private Const INPUT_HARDWARE = 2 'Hardwarenachricht

Private LV_Log As ListView
 
 
Public Property Get ERR_ACTIVATE_GETPATH_USER_HIT_CANCEL()
   ERR_ACTIVATE_GETPATH_USER_HIT_CANCEL = vbObjectError + 1
End Property

 
Private Sub SendKey(char As Byte, Optional Keyup = True)
   Dim IT As INPUT_TYPE
   
   With IT
      'KeyDown
       .dwType = INPUT_KEYBOARD
       .xi.wVk = char
       .xi.wScan = MapVirtualKey(char, 0)
       .xi.dwFlags = 0
       If SendInput(1&, IT, 28&) = 0 Then Debug.Print "Sendingkeys failed."
      
      'KeyUp
       .xi.dwFlags = IIf(Keyup, KEYEVENTF_KEYUP, 0)
       If SendInput(1&, IT, 28&) = 0 Then Debug.Print "Sendingkeys failed."
    End With
   
    Sleep (10)
   'Process Windows messages
'    DoEvents
      
End Sub

Private Sub SendKeys(vbKeys$)
   Dim i As Integer
   For i = 1 To Len(vbKeys)
      SendKey (Asc(Mid(vbKeys, i)))
   Next

End Sub



Public Sub Winhex_Activate()
 ' test if Hexworks is still running
   On Error Resume Next
   AppActivate "Hex Workshop"
   If Err <> 0 Then
    
    ' Start Hexworks
      Do
         Err.Clear
         ProcID = Shell(HWORKS_Path & " " & Quote(Filename), vbNormalFocus)
         If Err Then 'And (Chk_HexWork <> vbUnchecked)
            Dim tmpPath$
            tmpPath = InputBox("Please enter path+Filename to HWORKS32.exe", Err.Description, HWORKS_Path)
            If tmpPath = "" Then
               'Chk_HexWork.value = vbUnchecked
               On Error GoTo 0
               Err.Raise ERR_ACTIVATE_GETPATH_USER_HIT_CANCEL, "ClsSendToHexEditor::Activate", "Path to exe not found"
               
               Exit Sub
            Else
               HWORKS_Path = tmpPath
            End If
         Else: Exit Do
         End If
      Loop While True
    
    ' Wait till Hexworks is loaded
'      Sleep (500)
   End If

 ' switch to Hexworks
   AppActivate "Hex Workshop"

End Sub
Public Sub CloseHexWorkshop()
   On Error Resume Next
  
  'close open Hexworks
   Shell Environ("systemroot") & "\System32\Taskkill.exe /pid " & ProcID, vbHide

End Sub
Public Function LV_Log_GetOffset(ListView As ListView, Index)

   Dim item As MSComctlLib.ListItem
   Set item = ListView.ListItems(Index)
   
   On Error Resume Next
   
   Dim tryCount
   For tryCount = 0 To 20
   
      Set item = LV_Log.ListItems(item.Index + tryCount)
      LV_Log_GetOffset = FrmMain.LV_Log_Ext.ListSubItem(item, "pos")
      
      Dim NoData As Boolean
      NoData = (LV_Log_GetOffset = "")
      If NoData Then
       ' Select next
       '  Set LV_Log.SelectedItem = ListView.ListItems(item.Index + 1)
      Else
         LV_Log_GetOffset = OffToVal(LV_Log_GetOffset)
         Exit For

      End If
   Next
   

End Function

Public Sub Winhex_JumpToSelectedItem( _
   ListView As ListView, Form As Form, _
         Offset_DataStart, _
         Offset_CodeStart _
)
   Set LV_Log = ListView
   
   Dim item As MSComctlLib.ListItem
   Set item = LV_Log.SelectedItem
   
 ' on empty lines do nothing
 '  If "" = FrmMain.LV_Log_Ext.ListSubItem(item, "pos").Text Then Exit Sub
   
   
   
   Dim ClickedOffset
   ClickedOffset = LV_Log_GetOffset(ListView, item.Index)
   Set item = LV_Log.SelectedItem

   
   Dim objFasCmd As FasCommando
   Set objFasCmd = item.Tag
   
   
 ' calculate absolute offset in fasfile
 ' and care for Modul / ModuleBase
   Dim offset
   offset = ClickedOffset + IIf(objFasCmd.ModulId = 0, _
                                      Offset_DataStart, _
                                      Offset_CodeStart)
                                    
   Winhex_Activate
                                    
 
 ' send goto offset X to Hexworks
   VBA.SendKeys "{ESC}{F5}%b%d" & offset & "~", 1000
   
   
   
 ' Calculate length
   Dim nextItem, Length&
   If item.Index < ListView.ListItems.count Then
      nextItem = LV_Log_GetOffset(ListView, item.Index + 1)
      Length = nextItem - ClickedOffset
      
   End If
   
 ' Error correction
   If Length <= 0 Then Length = 1
   
   
   
  ' Keyinput: Shift Down - for an unknown reason Hex Workshop 4.10 "+{RIGHT}" don't work
    SendKey vbKeyShift, False
  
  ' Keyinput: SHIFT+RIGHT...
  ' For an unknown reason 'SendKey vbKeyShift, False' don't work Hex Workshop 2.54 but
  ' "+{RIGHT}" does
    VBA.SendKeys "+{RIGHT " & Length & "}", 1
  
  ' Keyinput: Shift UP
    SendKey vbKeyShift

 ' Activated Fas-decomp
   Form.SetFocus

   
End Sub


Private Sub Class_Initialize()
   
   HWORKS_Path = GetSetting(App.EXEName, "Hexworks", "Path", App.Path & "\Hex Workshop\HWORKS32.exe")

End Sub

Private Sub Class_Terminate()
   If HWORKS_Path <> "" Then SaveSetting App.EXEName, "Hexworks", "Path", HWORKS_Path
   
   CloseHexWorkshop
End Sub
