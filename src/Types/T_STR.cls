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
'   toText = Replace(m_value, vbLf, "\n")
'   toText = Replace(toText, "\", "\\")
'   toText = Replace(toText, """", "\""")

   toText = Quote(Escape(m_value))
End Property

Private Function Escape(ByRef Str)
   Dim inSize:   inSize = Len(Str)
   
   If inSize = 0 Then Exit Function
   
 ' init outbuffer
   Dim OutSize: OutSize = 0
   Escape = Space(inSize * 2) ' Assuming all are escape chars
   
   Dim i, o: o = 1
   For i = 1 To inSize
      
    ' 1. Get char
      Dim char_in
      char_in = Mid(Str, i, 1)
      
      Dim Changed As Boolean:    Changed = True
      Dim char_out:              char_out = char_in
      
    ' 2. Translate
      Select Case char_in
         Case vbLf:         char_out = "n"
         Case vbCr:         char_out = "r"
         Case vbTab:        char_out = "t"
         Case Chr(&H1B):    char_out = "e"
         
         Case """", "\":
         Case Else::           Changed = False
      End Select
    ' incase you also need to implement that  \000 octal thing
    ' also adjust the init
      
      
      ' Mask
      If Changed Then
         char_out = "\" & char_out:
      End If
    
    
    ' 3. Set char
      Mid(Escape, o) = char_out:   Inc o, Len(char_out)
      Inc OutSize, Len(char_out)
      
   Next
   
 ' shrink to real size
   Escape = Left(Escape, OutSize)
   
End Function

