Attribute VB_Name = "fas_file"
Option Explicit

'Assign each command some color
Public Function GetColor_Cmd(cmd)
  '
   GetColor_Cmd = RandInt_24bit(cmd)
End Function

Public Function GetColor_Type(TypeName)
  '
   GetColor_Type = "&h" & Right(ADLER32(TypeName), 6)
End Function


Private Function RandInt_24bit(Seed)
  'Note A negative init numbers for rnd() is considered as seed
   RandInt_24bit = Rnd(Seed * -1) * &HFFFFFF
End Function


'Public Function JoinASSym(ParamArray items())
'   Dim item
'   Dim tmp As New clsStrCat, tmp2$
'   For Each item In items(0)
'      tmp2 = item
'      If Left(tmp2, 1) <> "'" Then tmp.Concat "'"
'      tmp.Concat tmp2
'      tmp.Concat " "
'   Next
'   tmp.RemoveLast 2
'
''   JoinToText = Join(items(0)) ' tmp.value
'   JoinASSym = tmp.value
'
'End Function

Public Function JoinToText(ParamArray items())
'   Dim item
'   Dim tmp As New clsStrCat
'   For Each item In items(0)
'      tmp.Concat item.toText
'      tmp.Concat " "
'   Next
'   tmp.RemoveLast 2

   JoinToText = Join(items(0)) ' tmp.value
'   JoinToText = tmp.value

End Function
'Create Lisp Token from Keyword
' print, "hello world" -> "(print "hello world")"

Public Function TokenFull(keyword, ParamArray params())
Err.Clear
On Error Resume Next
   TokenFull = TokenOpen(keyword) & JoinToText(params(0)) & TokenClose(keyword)
   If Err = 13 Then 'Type mismatch
      Err.Clear
      TokenFull = TokenOpen(keyword) & JoinToText(params) & TokenClose(keyword)
      If Err Then
'         TokenFull = TokenOpen(keyword) & params(0).ToText & TokenClose(keyword)
         If Err Then Stop
      End If
   End If
   
End Function

Public Function TokenOpen(keyword, Optional IndentLevel = 0)
   TokenOpen = GetIndent(IndentLevel) & "(" & keyword & " "
End Function
Public Function TokenClose(Optional keyword)
   TokenClose = ")"
End Function

Public Function TokenRemove(Expr)
   TokenRemove = Mid(Expr, 1 + InStr(Expr, "("))
End Function


Public Function GetIndent(IndentLevel)
   GetIndent = Space(3 * IndentLevel)
End Function

