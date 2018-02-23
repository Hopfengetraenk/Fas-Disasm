VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "T_INT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Integers"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const opcode& = &H3B 'Ld_INT8
'(setq a 123 r 3.45 s "Hello!" x '(a b c))
'(setq f (open "name" "r"))
'(type a)                    returns  INT

'Public isNil As Boolean
Public isT As Boolean


Public value
Attribute value.VB_VarUserMemId = 0
Public Function toText()
   On Error Resume Next
 '  If isNil Then
 '     ToText = "nil"
 '  Else
   If isT Then
      toText = "T"
   Else
      toText = CStr(value)
   End If
End Function
