VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "T_SUBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Internal AutoLISP functions or functions loaded from compiled (FAS or VLX) files\r\n\r\nFunctions in LISP source files loaded from the AutoCAD Command prompt may also appear as SUBR"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Subroutine
'(type SETQ)                    returns  SUBR

Option Explicit
Public value
Attribute value.VB_VarUserMemId = 0

Public Property Get toText() As Variant
   toText = value
End Property
