Attribute VB_Name = "fas_file"
Option Explicit

'Assign each command some color
Public Function GetColor_Cmd(Cmd)
  '
   GetColor_Cmd = RandInt_24bit(Cmd)
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
'   Dim item, items_Count
'   Dim tmp As New clsStrCat
'   For Each item In items(0)
'      tmp.Concat item
'      If (items_Count And &H7) = &H7 Then
'         tmp.Concat vbCrLf
'      Else
'         tmp.Concat " "
'      End If
'
'      Inc items_Count
'   Next
'   tmp.RemoveLast 2

   JoinToText = Join(items(0)) ' tmp.value
 '  JoinToText = tmp.value

End Function
'Create Lisp Token from Keyword
' print, "hello world" -> "(print "hello world")"

Public Function TokenFull(Keyword, ParamArray Params())
Err.Clear
On Error Resume Next
   TokenFull = TokenOpen(Keyword) & JoinToText(Params) & TokenClose(Keyword)
   TokenFull = TokenOpen(Keyword) & JoinToText(Params(0)) & TokenClose(Keyword)
   If TokenFull = "" Then Stop
   
   Err.Clear
End Function

Public Function TokenComment(Line, Optional IndentLevel = 0)
   If Line <> "" Then
      TokenComment = GetIndent(IndentLevel) & ";;; " & Line
   End If
End Function

' ((= (atof (getvar 'AcadVer)) 18.0) ;| 2010 code here |;)
Public Function TokenInlineComment(Text)
   TokenInlineComment = ";| " & Text & " |;"
End Function


Public Function TokenOpen(Keyword, Optional IndentLevel = 0)
   TokenOpen = GetIndent(IndentLevel) & "(" & Keyword & " "
End Function
Public Function TokenClose(Optional Keyword, Optional IndentLevel = 0)
   TokenClose = GetIndent(IndentLevel) & ")"
End Function

Public Function TokenRemove(Expr)
   TokenRemove = Mid(Expr, 1 + InStr(Expr, "("))
End Function


Public Function GetIndent(IndentLevel)
   GetIndent = Space(3 * IndentLevel)
End Function

