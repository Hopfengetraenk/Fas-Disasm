VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "T_SYM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "String"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'(setq a 123 r 3.45 s "Hello!" x '(a b c))
'(setq f (open "name" "r"))
'(type 'a)                   returns  SYM
Private m_value
Attribute m_value.VB_VarUserMemId = 0
Const opcode_fas& = &H5B 'ld_SYM
Const opcode_fsl = &H56

Private knownSymbols As New Collection

Public AsSYM As Boolean

Public Property Get value()
Attribute value.VB_UserMemId = 0

   Err.Clear

   If AsSYM And Collection_IsAlreadyIn(knownSymbols, m_value) Then
      value = "'" & m_value
   Else
      value = m_value
   End If
            
If Err Then Stop
End Property

Public Property Let value(ByVal bNewValue)

   m_value = bNewValue

End Property

Public Property Set value(ByVal bNewValue)
Stop
   Set m_value = bNewValue

End Property


Public Property Get toText() As Variant
   
   toText = m_value
End Property
            
Private Sub Class_Initialize()
   Dim item
   For Each item In Split("EXRXSUBR LIST SUBR FILE ENAME PICKSET REAL INT STR")
      knownSymbols.add "", item
   Next
End Sub
