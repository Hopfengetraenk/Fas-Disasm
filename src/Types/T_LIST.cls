VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "T_LIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "List"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const opcode& = &H39 'Ld_LIST
'(setq a 123 r 3.45 s "Hello!" x '(a b c))
'(setq f (open "name" "r"))
'(type x)                    returns  LIST

Public m_items As Collection
Public Output As clsStrCat

Private m_NestLevel&

Const List_PreFix$ = "("
Const List_InFix$ = " "
Const List_PostFix$ = ")"
'TokenFull("",

Private m_bvalue
Public isCons As Boolean

Public AsSYM As Boolean

Public Property Get data(idx&)
Err.Clear

   data = m_bvalue(idx)
If Err Then Stop
End Property

Function GetFormated()
   GetFormated = TokenFull( _
            IIf(isCons, "cons", ""), _
            m_bvalue)
End Function
Public Property Get value()
Attribute value.VB_UserMemId = 0
            
   value = GetFormated()
            
   If AsSYM Then value = "'" & value
            
'   value =  m_bvalue
If Err = 13 Then
   value = TokenFull("", m_bvalue)
End If

If Err Then Stop
End Property

Public Property Let value(ByVal bNewValue)

   m_bvalue = bNewValue

End Property

Public Property Set value(ByVal bNewValue)

   Set m_bvalue = bNewValue

End Property




Public Function add(item)
   m_items.add item
End Function

Public Function toText()

   toText = "'" & GetFormated()
'   On Error Resume Next
'   toText = "T_LIST-ToText"

'   m_NestLevel = 0
'
'   Set Output = New clsStrCat
'   Dim item
'   For Each item In m_bvalue
'      ToTextRek item
'   Next
'
'   Output.RemoveLast Len(List_InFix)
'
'   ToTextRek m_bvalue
'
'   ToText = Output.value
'
End Function


Public Sub ToTextRek(item)
   Inc m_NestLevel
   
   Output.Concat List_PreFix
   
      If TypeOf item Is Collection Then
         ToTextRek item
      Else
         Output.Concat item & List_InFix
      End If
   
   Output.Concat List_PostFix
   
End Sub


Public Property Get NestLevel() As Variant
   NestLevel = m_NestLevel
End Property

