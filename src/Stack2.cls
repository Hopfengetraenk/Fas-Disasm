VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stack2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Const STACK_MIN = 0
'Private Storage As New Collection
Private Storage()
Public ESP As Long

Public Property Let wholeStack(ByVal vdata As Variant)
   Storage = vdata
   ESP = UBound(Storage)
End Property

Public Property Get wholeStack() As Variant
    wholeStack = Storage
End Property




'Public Property Let data(value&)
'End Property
'
'Public Property Let data(value&)
'End Property



Public Property Let size(value&)
   ReDim Preserve Storage(value)
End Property

Public Property Get size&()
    size = UBound(Storage)
End Property



Public Sub push(data As Variant)
'Try
 On Error GoTo push_err
    Storage(ESP) = data
    ESP = ESP + 1
 Err.clear
'catch
push_err:
 Select Case Err
    Case 0:
    Case 9: 'Index außerhalb des gültigen Bereichs
        ReDim Preserve Storage(ESP + 10)
        If ESP >= LBound(Storage) Then Resume
    Case Else
        Err.Raise Err, , Err.Description
End Select
'Finally


End Sub

Public Property Get data() As Long
Attribute data.VB_UserMemId = 0
   PreviewPop = Storage(ESP - 1)
End Property

Public Property Let data(vdata As Long)
   Storage(Me.ESP - 1) = vdata
End Property

Public Function pop() As Variant
   'Try
    On Error GoTo pop_err
    ESP = ESP - 1
    pop = Storage(ESP)

    Err.clear
'catch
pop_err:
Select Case Err
    Case 0:
    Case 9: 'Index außerhalb des gültigen Bereichs
        If ESP <= STACK_MIN Then ESP = STACK_MIN
'->note for developer
' rightclick  Select "switch" -> "don't stop when errors"
        Err.Raise vbObjectError + 1001, , "Stack is empty - Pop is not possible."
    Case Else
        Err.Raise Err, , Err.Description
End Select
'Finally

End Function

Public Function popArray(NumberOfElements) As Variant
    
   If NumberOfElements < 1 Then Exit Function
   Dim tmp
   ReDim tmp(1 To NumberOfElements)
      
   Dim i
   For i = UBound(tmp) To LBound(tmp) Step -1
      tmp(i) = Me.pop
   Next
   
   popArray = tmp
End Function

Public Function clone() As Stack
   Set clone = New Stack
   clone.wholeStack = Me.wholeStack
   clone.ESP = Me.ESP
End Function

Private Sub Class_Initialize()
   push 0
End Sub
