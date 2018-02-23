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
Public isLambda As Boolean
'      ^- it's a Lambda function def when it's invoked via opcode &H35_ld_USUBR


Public Startoffset As Variant

Public ModulId

Public ModulStream As FileStream


Public Endoffset As Variant

Public Args As New Collection

Public Vars

Public VarsCount


Private m_LocalVars

Private mvarNumOfArgs As Long
Public Property Let NumOfArgs(ByVal vData As Long)
    mvarNumOfArgs = vData
End Property


Public Property Get NumOfArgs() As Long
    NumOfArgs = mvarNumOfArgs
End Property



'MakeLocalVars
Public Sub MakeLVars(VarsCount, Optional PreFixForUninit = "_fas")
     
     VarsCount = VarsCount + 1
     ReDim m_LocalVars(VarsCount)

     
   ' Make LocalVars
     Dim i
     For i = 0 To VarsCount
        Dim tmp As Fas_LVar
        Set tmp = New Fas_LVar
        tmp.value = "" & PreFixForUninit & i
        
        Set m_LocalVars(i) = tmp
     Next
     
End Sub

Public Property Get LocalVars(idx) As Variant
   Set LocalVars = m_LocalVars(idx)
End Property

Public Property Let LocalVars(idx, ByVal vNewValue As Variant)
    m_LocalVars(idx).value = vNewValue
End Property
