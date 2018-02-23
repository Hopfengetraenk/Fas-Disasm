VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FasLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public m_items As Collection
Public Output As clsStrCat

Private m_NestLevel&

Const List_PreFix$ = "("
Const List_InFix$ = " "
Const List_PostFix$ = ")"



'Public Function create(item)
   
'End Function

Public Function add(item)
   m_items.add item
End Function

Public Function ToText()
   
   m_NestLevel = 0
   
   Output = New clsStrCat
   Dim item
   For Each item In m_items
      ToTextRek item
   Next
   
   Output.RemoveLast Len(List_InFix)

   ToTextRek m_items
   
   ToText = Output.value

End Function


Public Sub ToTextRek(item)
   Inc m_NestLevel
   
   Output.Concat List_PreFix
   
      If TypeOf item Is Collection Then
         ToTextRek item
      Else
         Output.Concat item & List_InFix
      End If
   
   Output.Concat List_PostFix
   
End Sub


Public Property Get NestLevel() As Variant
   NestLevel = m_NestLevel
End Property

