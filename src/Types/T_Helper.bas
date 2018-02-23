Attribute VB_Name = "T_Helper"
Option Explicit


Function make_NIL() As T_NIL
   Set make_NIL = New T_NIL
End Function

Function make_INT(value) As T_INT
   Set make_INT = New T_INT
   make_INT = value
End Function

Function make_REAL(value As String) As T_REAL
   Set make_REAL = New T_REAL
   make_REAL = value
End Function

Function make_STR(value) As T_STR
   Set make_STR = New T_STR
   make_STR = value
End Function
Function make_SYM(value) As T_SYM
   Set make_SYM = New T_SYM
   make_SYM = value
End Function

Function make_LIST(value) As T_LIST
   Set make_LIST = New T_LIST
   make_LIST.value = value
End Function

Function make_USUBR(value As String) As T_USUBR
   Set make_USUBR = New T_USUBR
   make_USUBR = value
End Function


Function make_ITEM(value) As E_ITEM
   Set make_ITEM = New E_ITEM
   make_ITEM = value
End Function
