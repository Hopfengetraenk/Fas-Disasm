VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Private Const STACK_MIN = 0
Public Storage As New Collection
'Private Storage()
Public ESP As Long
Attribute ESP.VB_VarUserMemId = 0
'
'Public Property Let wholeStack(ByVal vData As Variant)
'   Storage = vData
'   ESP = UBound(Storage)
'End Property
'
'Public Property Get wholeStack() As Variant
'    wholeStack = Storage
'End Property

Public Property Let Current(ByVal vData As String)
    On Error GoTo Err_Current
      If ESP > 0 Then Me.pop
      Me.push vData
    Err.clear
Err_Current:
    Select Case Err
      Case 0
      Case vbObjectError + 1001
         Resume Next
    End Select
End Property


'Public Property Get Current() As String
 '   Stop
    'Current = Storage(ESP)
'End Property







'
'Public Property Let size(value&)
'   ReDim Preserve Storage(value)
'End Property
'
'Public Property Get size&()
'    size = UBound(Storage)
'End Property



Public Sub push(data As Variant)
 
'   On Error GoTo push_err

   If (Storage.count = 0) Or (ESP >= Storage.count) Then
     'Insert into empty list or insert at the End
      Do While (ESP > Storage.count)
        Storage.Add "FillData"
      Loop
      Storage.Add data
   Else
      Storage.Add data, , , ESP + 1
      Storage.Remove (ESP + 1) 'delete element
      
   End If
   ESP = ESP + 1
 
'catch
'push_err:
' Select Case Err
'    Case 0:
'    Case 5:
''      If ESP = 0 Then
'         Storage.Add data
'         Resume Next
'   '   Else
'  '       Stop
' '     End If
'    Case 9: 'Index außerhalb des gültigen Bereichs
'      If ESP > 0 Then
'        Storage.Add "FillData"
'      ElseIf ESP = 0 Then
'         Storage.Add data, , 1
'         Resume Next
'      Else
'         ESP = 0
'      End If
'       Resume
'    Case Else
'        Err.Raise Err, , Err.Description
'End Select
''Finally


End Sub

Public Function PreviewPop() As Variant
   PreviewPop = Storage(ESP)
End Function

Public Function pop() As Variant
   'Try
    On Error GoTo pop_err
    pop = Storage(ESP)
    ESP = ESP - 1

    Err.clear
'catch
pop_err:
Select Case Err
    Case 0:
    Case 9: 'Index außerhalb des gültigen Bereichs
'        If ESP <= STACK_MIN Then ESP = STACK_MIN
'->note for developer
' rightclick  Select "switch" -> "don't stop when errors"
'        qw
        Err.Raise vbObjectError + 1001, , "Stack is empty - Pop is not possible."
    Case Else
        Err.Raise Err, , Err.Description
End Select


'Finally

End Function

Public Function popArray(NumberOfElements) As Variant
   
   Dim tmp
   If NumberOfElements < 1 Then
      ReDim tmp(1 To 1)
   
   Else
      ReDim tmp(1 To NumberOfElements)
         
      Dim i
      For i = UBound(tmp) To LBound(tmp) Step -1
         tmp(i) = Me.pop
      Next
   End If
   
   popArray = tmp
   
End Function

'Public Function clone() As Stack
'   Set clone = New Stack
'   clone.wholeStack = Me.wholeStack
'   clone.ESP = Me.ESP
'End Function

