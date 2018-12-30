VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "T_NIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const opcode& = &H1  'Ld_NIL
'(type NIL)                   returns  NIL

Public SupressOutput As Boolean
'Quote(value)


Public Property Get value() As Variant
Attribute value.VB_UserMemId = 0
   If SupressOutput = False Then
      value = "NIL"
   '   value = "'()"
   End If
End Property


Public Property Get toText() As Variant
   toText = value
End Property
