VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "T_STR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const opcode& = &H55 'Load_STR
'(setq a 123 r 3.45 s "Hello!" x '(a b c))
'(setq f (open "name" "r"))
'(type 's)                   returns  STR
Private m_value
Attribute m_value.VB_VarUserMemId = 0

'Quote(value)

Public Property Get value() As Variant
Attribute value.VB_UserMemId = 0
   value = toText() 'm_value
End Property

Public Property Let value(ByVal vNewValue As Variant)
   m_value = vNewValue
End Property


Public Property Get toText() As Variant
   toText = Quote(m_value)
End Property

