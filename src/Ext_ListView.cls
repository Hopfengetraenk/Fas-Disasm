VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ext_ListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mListView As ListView
Private Const VisibleLinesAbove& = 8



' helper for the ListView control
'   Uses Column.key to access a subItem
'     of the given ListItem
Public Property Get ListSubItem(li As ListItem, Key) As ListSubItem

   Debug.Assert Not (mListView Is Nothing)
   
   With mListView
   
    ' Get Column Index
      Dim SubListID%
      SubListID = .ColumnHeaders(Key).Index - 1
      '^- '.Index-1' since item counting is
      ' item_1, subitem_1, subitem_2, subitem_2...   while colum counting is
      '  col_1,     col_2,     col_3,...
      
    ' Redim SubItems
      li.SubItems(SubListID) = li.SubItems(SubListID)
      
    ' get ListSubItem object
      Set ListSubItem = li.ListSubItems(SubListID)
      
   End With
End Property

' Why this: the item show by Listview::EnsureVisible just sticks at the top or bottom
'   This EnsureVisible will seek more in the middle
Public Sub EnsureVisible(li As MSComctlLib.ListItem)
   
   Debug.Assert Not (mListView Is Nothing)
   
   Dim item_Index
   item_Index = li.Index
   With mListView.ListItems
            
'          ' Jump to target
'          ' 1.to the end
'            .ListItems(.ListItems.count).EnsureVisible
'          ' 2.to the item
'            item.EnsureVisible
'          ' 3.Scroll up some lines
'            .ListItems(item.Index - NAV_SCROLLDOWN_LINES).EnsureVisible

      
    ' 1. to Start
     .item(.count).EnsureVisible
      
      
    ' 2. to some items further as the target
      Dim item_index_down
      item_index_down = item_Index - VisibleLinesAbove
      If item_index_down > 0 Then
         .item(item_index_down).EnsureVisible
      'Else
         'skip this step - since that'll be beyond the end of the item list
      End If
      
    ' 3.to target item
      .item(item_Index).EnsureVisible
      
   End With
   
   DoEvents
   
End Sub


Public Sub Create(ListView As ListView)
   Set mListView = ListView
End Sub


'_________________________________________________________________________________________


'Encode_offset means an Offset created with OffToStr()
' Example OffToStr(1022) -> "$FFE"
' TODO: use ClsOffset instead of String
Public Static Function offset_enc(offset, ModulId)
   offset_enc = "off:" & offset & "_" & ModulId
'   Debug.Print offset_enc
End Function
'Public Static Function offset_dec(offset_enc, ModulID)
'   offset_dec = Split(offset_enc, "off:")(2)
'   ' & "_" & ModulID
'End Function

Public Static Property Get OffsetKeyGet(ModulId, offset) As ListItem
'li As ListItem
'   Dim FasCommando As FasCommando
'   Set FasCommando = li.Tag

   Set OffsetKeyGet = mListView.ListItems(offset_enc(offset, ModulId))
End Property

Public Static Property Let OffsetKey(li As ListItem, ModulId, offset)
'
'   Dim FasCommando As FasCommando
'   Set FasCommando = li.Tag
   li.Key = offset_enc(offset, ModulId)

End Property




