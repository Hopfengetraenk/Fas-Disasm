VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FasFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public FuncName As Variant
Attribute FuncName.VB_VarUserMemId = 0

Public Startoffset As Variant

Public Endoffset As Variant

Public Args As Collection

'lokale Variable(n) zum Zuweisen der Eigenschaft(en)
Private mvarNumOfArgs As Long 'lokale Kopie
Public Property Let NumOfArgs(ByVal vData As Long)
'wird beim Zuweisen eines Werts zu der Eigenschaft auf der linken Seite einer Zuweisung verwendet.
'Syntax: X.NumOfArgs = 5
    mvarNumOfArgs = vData
End Property


Public Property Get NumOfArgs() As Long
'wird beim Ermitteln eines Eigenschaftswertes auf der rechten Seite einer Zuweisung verwendet.
'Syntax: Debug.Print X.NumOfArgs
    NumOfArgs = mvarNumOfArgs
End Property



