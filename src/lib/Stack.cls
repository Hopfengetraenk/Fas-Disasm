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
Public esp As Long
Attribute esp.VB_VarUserMemId = 0
'
'Public Property Let wholeStack(ByVal vData As Variant)
'   Storage = vData
'   ESP = UBound(Storage)
'End Property
'
'Public Property Get wholeStack() As Variant
'    wholeStack = Storage
'End Property

Public Property Let Current(ByVal vData) ' As String)
    On Error GoTo Err_Current
      If esp > 0 Then Me.popIntoVoid
      Me.push vData
    Err.Clear
Err_Current:
    Select Case Err
      Case 0
      Case vbObjectError + 1001
         Resume Next
    End Select
End Property


Public Property Get Current() ' As String
    On Error Resume Next
    
    Fetch Current
    
    If Err = 9 Or Err = 5 Then '9 Subscription out of Range / 5 inv arg
      Dim Errmsg$
      
      If esp <= 0 Then
         Errmsg = "Stack is empty.   Requested SP: " & esp
      Else
         Errmsg = "Stackpointer esp=" & esp & " out of range."
       End If
       
       On Error GoTo 0
       Err.Raise ERR_STACK_IS_EMPTY, "Stack::Current", Errmsg
    End If
    
End Property







'
'Public Property Let size(value&)
'   ReDim Preserve Storage(value)
'End Property
'
'Public Property Get size&()
'    size = UBound(Storage)
'End Property



Public Sub push(data As Variant)
 
   On Error GoTo push_err

   If (Storage.count = 0) Or (esp >= Storage.count) Then
     'Insert into empty list or insert at the End
      Do While (esp > Storage.count)
        Storage.add "FillData"
      Loop
      Storage.add data
   Else
      Storage.add data, , , esp + 1
      Storage.Remove (esp + 1) 'delete element
      
   End If
   Inc esp
 
'catch
push_err:
 Select Case Err
    Case 0:
    Case 5:
'      If ESP = 0 Then
         Storage.add data
         Resume Next
   '   Else
  '       Stop
 '     End If
    Case 9: 'Index außerhalb des gültigen Bereichs
      If esp > 0 Then
        Storage.add "FillData"
      ElseIf esp = 0 Then
         Storage.add data, , 1
         Resume Next
      Else
         esp = 0
      End If
       Resume
    Case Else
        Err.Raise Err, , Err.Description
End Select
'Finally


End Sub


Sub popIntoVoid()
    esp = esp - 1
End Sub

Private Sub Fetch(ByRef Out)
   If IsObject(Storage(esp)) Then
      Set Out = Storage(esp)
    Else
      Out = Storage(esp)
   End If

End Sub

Property Get pop() As Variant
   'Try
    On Error GoTo pop_err
    
    Fetch pop

    esp = esp - 1

    Err.Clear
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

End Property

Property Get popArray(NumberOfElements) As Variant
   
   Dim tmp
   If NumberOfElements < 0 Then
      Stop 'Error NumberOfElements is negative

   ElseIf NumberOfElements = 0 Then
      tmp = Array()
   
   Else
      ReDim tmp(1 To NumberOfElements)
         
      Dim i
      For i = UBound(tmp) To LBound(tmp) Step -1
         If IsObject(Me.Current) Then
            Set tmp(i) = Me.Current
         Else
            Let tmp(i) = Me.Current
         End If
         Me.popIntoVoid
      Next
   End If
   
   popArray = tmp
   
End Property

'Public Function clone() As Stack
'   Set clone = New Stack
'   clone.wholeStack = Me.wholeStack
'   clone.ESP = Me.ESP
'End Function

