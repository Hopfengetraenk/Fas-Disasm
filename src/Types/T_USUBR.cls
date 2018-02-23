VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "T_USUBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "User-defined functions loaded from LISP source files"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Const opcode& = &H14 'Ld_USUBR
' User Subroutine
' _$ (DEFUN one() 1)
' _$ (TYPE  one)                    returns  SUBR
' USUBR
Option Explicit


Public Name

'Public isLambda As Boolean
''      ^- it's a Lambda function def when it's invoked via opcode &H35_ld_USUBR

Public Start
'Public params
Private mvarvalue As Variant 'lokale Kopie

Public Property Let value(ByVal vData As Variant)
    mvarvalue = vData
End Property


Public Property Set value(ByVal vData As Variant)
    Set mvarvalue = vData
End Property


Public Property Get value() As Variant
Attribute value.VB_UserMemId = 0
    If IsObject(mvarvalue) Then
        Set value = mvarvalue
    Else
        value = mvarvalue
    End If
End Property



Public Property Get toText() As Variant
   toText = value
End Property
