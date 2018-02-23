VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "T_PICKSET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Selection sets"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'The PICKFIRST data type:  a selection set, commonly based on the objects in a drawing that the user wants to modify or interact with
'; Selects the last object drawn at 0,0,0
'(SETQ ssPt (ssget '(0 0 0)))
'(TYPE ssPt)
'Give PICKFIRST

Option Explicit

Public value
Attribute value.VB_VarUserMemId = 0
