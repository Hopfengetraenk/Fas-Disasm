VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fas_Brancher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const SkipInternalData As Boolean = 0

Private Const PartsSeperatorOnTheLine As String = " "
Private Const PartsSeperator As String = PartsSeperatorOnTheLine & vbCrLf

Public isWhile As Boolean
Public isIf_noRet As Boolean
Public isRepeat As Boolean
Public isForeach As Boolean
Public isVlaxFor As Boolean
Public isIf As Boolean

Public Keyword

Public FileStream As FileStream
Public FasStack As Stack
Public level

      
Dim Target As Collection


Public FasFile As FasFile


Public ForEach_item As Object
Public VlaxFor_item As Object

   
Public Brancher_Cond As New Brancher_Cond



Public Property Get Goto_Size() As Long
   If FasFile.FasFile_Version = 4 Then
      Goto_Size = 1 + 4 'Size for 0x57 Goto opCode
   Else
      Goto_Size = 1 + 2 'Size for 0x0d Goto opCode
   End If
End Property



Public Sub create(FileStream As FileStream, FasStack As Stack, FasFile As FasFile, level)
   Set Me.FileStream = FileStream
   Set Me.FasStack = FasStack
   Me.level = level
   
   Set Me.FasFile = FasFile
   
   Set Target = New Collection
   
   Brancher_Cond.create Me, FileStream


End Sub

Public Function Repeat_PatchFor_h4B(Cmd, Param2) As Boolean
' check for '(<= <xxx> 0)'
  If (Cmd = &H4B) And _
     (Param2 = 0) Then
     
   '
     Dim NextCmd
     NextCmd = FileStream.int8: FileStream.Move -1
     
     Repeat_PatchFor_h4B = (NextCmd = &H67)

     
  End If
' without this patch repeat will decompiled like this
' "repeat ((<= (LENGTH A) 0))"
End Function


'Must be called after &H67: 'Branch_32 is true
Public Function GetKeyWord(Last_FasCmdline As FasCommando, FasCmdline As FasCommando)

   Debug.Assert (FasCmdline.Commando = &H67) Or (FasCmdline.Commando = &HD)  'That's the referencepoint

   With FileStream
     '.StorePos
      Dim FileStream_oldPos
      FileStream_oldPos = .Position
      
     'seek to cmd_&H67 GotoLocation
      Dim GotoDelta
      GotoDelta = FasCmdline.Parameters(1)
      
      .Move GotoDelta - Goto_Size
      
      Dim tmp
      .Move -1: tmp = .int8
      
     'Get GotoAdress
      Dim Goto_Addr&
      Goto_0x57 Goto_Addr
      
      
   '... while  -> when there's at $p1-(1+4) some Goto(&H57) that jumps backwards
      
      isWhile = Goto_Addr < 0
      
      
      'isIf_noRet = Goto_Addr = 0
      isIf_noRet = (tmp = &HA) Or (Goto_Addr = 0)
      
      Debug.Print "Brancher Heuri", OffToStr(FasCmdline.Position) & " " & Goto_Addr
      
      ' is it Repeat?
       
       isRepeat = Last_FasCmdline.Commando = &H4B '(<= a b)
      
      If isRepeat Then
         Keyword = "repeat"
         
      ElseIf isWhile Then
         
         isForeach = isClearArgsNVars_0x19(FileStream)
         
         isVlaxFor = isGetVar_0x5c(FileStream)
         
         If isForeach Then
            Keyword = "foreach" & PartsSeperatorOnTheLine & ForEach_item:    Set ForEach_item = Nothing
            
         ElseIf isVlaxFor Then
            Keyword = "vlax-for" & PartsSeperatorOnTheLine & ForEach_item:   Set ForEach_item = Nothing
         Else
            Keyword = "while"
         End If
         
'      ElseIf isIf2 Then
'         Keyword = "if"
      
      Else
         Keyword = "if"
         isIf = True
      End If
      
      
'      .RestorePos
      .Position = FileStream_oldPos

   End With
   
   GetKeyWord = Keyword
End Function

Public Function Goto_0x57_isNullJump(Optional FromEnd = False) As Boolean
   With FileStream
   
      If FromEnd Then .Move -Goto_Size
      
      Dim GotoPos&
      Goto_0x57 GotoPos
      Goto_0x57_isNullJump = GotoPos = 0 'isNull
      'Debug.Assert Goto_0x57_isNullJump
   
     'seek back to begin
      If Not FromEnd Then .Move -Goto_Size
   End With
End Function

Public Function Goto_0x57_forward(Optional DoGoto = True, Optional FromEnd = False)
   Goto_0x57_forward = Goto_0x57_seek(True, DoGoto, FromEnd)
End Function
Public Function Goto_0x57_backward(Optional DoGoto = True, Optional FromEnd = False)
   Goto_0x57_backward = Goto_0x57_seek(False, DoGoto, FromEnd)
End Function
Private Function Goto_0x57_seek(isforward As Boolean, Optional DoGoto = True, Optional FromEnd = False)
   With FileStream

      If FromEnd Then .Move -Goto_Size
      
      Dim GotoPos&
      Goto_0x57 GotoPos
'      Debug.Assert GotoPos <> 0 'use Goto_0x57_isNullJump to check for this
      If isforward Then
         Debug.Assert GotoPos >= 0 'isForwards
      Else
         Debug.Assert GotoPos < 0 'isBackwards
      End If
   
      Goto_0x57_seek = GotoPos + .Position
      
      If DoGoto Then .Position = Goto_0x57_seek
      
'      If FromEnd Then .Move -Goto_Size

   End With
End Function

' Check is current location is Goto and destination borders
' returns its destination
Private Function Goto_0x57(ByRef Goto_Addr__OUT As Long) As Boolean
    With FileStream
         Dim Goto_cmd As Byte
         Goto_cmd = .int8
         
         'Dim is_Goto_0x57 As Boolean
         If FasFile.FasFile_Version = 4 Then
            Goto_0x57 = Goto_cmd = &H57
         Else
            Goto_0x57 = Goto_cmd = &HF
         End If
         
       ' oops probably wrong location - this is no goto
         Debug.Assert Goto_0x57 = True
         
         Goto_Addr__OUT = GotoTargetGet
      
    End With
End Function
Private Function GotoTargetGet() As Long

   With FileStream
      If FasFile.FasFile_Version = 4 Then
         GotoTargetGet = .int32Sig
      Else
         GotoTargetGet = .int16
      End If
    ' Goto should seek somewhere inside the file
      Debug.Assert RangeCheck(GotoTargetGet + .Position, .Length, 0)
   End With
End Function
      
Private Function GotoTarget() As Long
      
      GotoTarget = GotoTargetGet()
     FileStream.Move GotoTarget

   
End Function

'General get 32 bit pointer (including zero in border check)
Public Function GotoTarget_Forward()
   GotoTarget_Forward = GotoTarget
   Debug.Assert GotoTarget_Forward >= 0
End Function
Public Function GotoTarget_Backward()
   GotoTarget_Forward = GotoTarget
   Debug.Assert GotoTarget_Forward <= 0
End Function




'Check is current location is Goto and returns its destination
Private Function isClearArgsNVars_0x19(FileStream) As Boolean
    With FileStream
         Dim NextCmd As Byte
         NextCmd = .int8
         .Move -1
         
          isClearArgsNVars_0x19 = NextCmd = &H19
         
    End With
End Function
'Check is current location is Goto and returns its destination
Private Function isGetVar_0x5c(FileStream) As Boolean
    With FileStream
         Dim NextCmd As Byte
         NextCmd = .int8
         .Move -1
         
          isGetVar_0x5c = NextCmd = &H5C
         
    End With
End Function

Public Sub Decompile(Last_FasCmdline As FasCommando, FasCmdline As FasCommando, Branch_Firstpart)
   Debug.Assert (FasCmdline.Commando = &H67) Or (FasCmdline.Commando = &HD) 'That's the referencepoint
   
   
   If isIf Then 'Or isIf_noRet
      DoIf Last_FasCmdline, FasCmdline, Branch_Firstpart
      
   ElseIf isForeach Then
      DoForEach Last_FasCmdline, FasCmdline, Branch_Firstpart
      
   ElseIf isVlaxFor Then
      DoVlaxFor Last_FasCmdline, FasCmdline, Branch_Firstpart
   
   ElseIf isRepeat Then
      
'      FasStack.Current = ""
'      FasStack.popIntoVoid
      
      Dim Repeat_Body
      Repeat_Body = DoBlock(FileStream.Position + FasCmdline.Parameters(1), level + 1)
      Repeat_Body = Join(Repeat_Body, PartsSeperator)
      FasStack.Current = PartsSeperator & Branch_Firstpart & PartsSeperator & _
                         Repeat_Body & TokenClose(Keyword)
                           

      
   ElseIf isWhile Then
      
      'FasStack.Current = ""
      
      Dim While_Body
      While_Body = DoBlock(FileStream.Position + FasCmdline.Parameters(1), level + 1)
      While_Body = Join(While_Body, PartsSeperator)
'      FasStack.popIntoVoid
      
      Dim While_Full
      While_Full = PartsSeperator & Branch_Firstpart & PartsSeperator & _
                         While_Body
                           
      While_Full = make_ITEM(While_Full & PartsSeperator & TokenClose(Keyword))
      
      FasFile.outp While_Full
      
   Else
      ' process  'Then' Branch
  '    FasStack.Current = Branch_Firstpart
      
 '     FasFile.InterpretStream_rek FileStream, Target, FileStream.Position + FasCmdline.Parameters(1) - 1, level + 1 ' , IIf(isRepeat, -1, 0)
      
      
      ' Close Branch/Loop sequence block
      'FasCmdline.Interpreted = FasCmdline.Interpreted & ")"
  '    FasStack.Current = TokenClose(Keyword)
   End If

End Sub


'   No RetValue           With RetValue
'   18    0 iArgs         18    0 iArgs
'                          1    1 ld NIL
'         ...                   ...
'   67    1 vlax-for N    67    1 vlax-for N
'                          A    0 Pop
'
'    3    1 VALUE          3    1 VALUE
'   35    1 ld_USUBR      35    1 ld_USUBR

'    A    0 Pop
'   57    1 GOTO          57    1 GOTO
Private Sub DoVlaxFor(Last_FasCmdline As FasCommando, FasCmdline As FasCommando, VlaxFor_Expr)

   On Error GoTo DoVlaxFor_err

   Dim EndPos
   EndPos = FileStream.Position + FasCmdline.Parameters(1) - 1
   
   
   
   Dim hasRetValue As Boolean
   hasRetValue = GetHasRetValue()
   
   
   Dim VlaxForBody
   VlaxForBody = DoProgn(EndPos, level + 1)
 '  Dim hasRetValue As Boolean
   hasRetValue = Last_FasCmdline.Commando <> &HA 'pop
   
   Dim VlaxFor_Full
   VlaxFor_Full = Join(Array("", _
                  GetIndent(-hasRetValue) & VlaxFor_Expr, _
                  VlaxForBody, _
                  GetIndent(-hasRetValue) & TokenClose, "" _
                  ), PartsSeperator)
   
   
   FasFile.outp make_ITEM(VlaxFor_Full)
   
   
   isVlaxFor = False
   Set VlaxFor_Expr = Nothing
   
Exit Sub
DoVlaxFor_err:
   Stop
   Resume
End Sub

Public Function GetHasRetValue() As Boolean

   Dim NextCmd
   NextCmd = FileStream.int8: FileStream.Move -1
   
   GetHasRetValue = (NextCmd = &HA)
   
  'just for checking structure
   If GetHasRetValue Then
     'just for checking structure
      Debug.Assert TypeOf FasStack.Current Is T_NIL
      FasStack.Current = make_ITEM("")
   End If

End Function
Public Sub Skip_Pop_0a()
   '  A                   pop dummy (decrease stack)
      Dim dummy0: dummy0 = FileStream.int8
      If dummy0 <> &HA Then
         Stop
         FileStream.Move -1
      Else
         dummy0 = FasStack.pop
      End If
End Sub
Public Sub Skip_Pop_0a__NoPop()
   '  A                   pop dummy (decrease stack)
      Dim dummy0: dummy0 = FileStream.int8
      If dummy0 <> &HA Then
         Stop
         FileStream.Move -1
      Else
'         dummy0 = FasStack.pop
      End If
End Sub




Private Sub DoForEach(Last_FasCmdline As FasCommando, FasCmdline As FasCommando, ForEach_Expr)
   On Error GoTo DoForEach_err

   Dim EndPos
   EndPos = FileStream.Position + FasCmdline.Parameters(1) - 1
   
   With FileStream
'      .StorePos

   Dim hasRetValue As Boolean
   hasRetValue = GetHasRetValue()

   If SkipInternalData Then
      '  00187 67    1C            If (FuncArg[0]==0)[Far] jump to 220          1
      '  00192 A                   pop dummy (decrease stack)                   0
       Skip_Pop_0a
      
      
      
      '
      '  Load current list element into N
      '  00193 5C    0             Push FuncArg[0]                              1
      '  00196 28                  push list element from (FuncArg[0])          1
      '  00197 6     3             setq N list element from (FuncArg[0])        0
         Debug.Assert .int8 = &H5C:   Dim dummy0: dummy0 = .int16
         Debug.Assert .int8 = &H28
         Debug.Assert .int8 = &H6:   Dim dummy1: dummy1 = .int16
      
      '
      '  Move FuncArg[0] to next list element
      '  00200 5C    0             Push FuncArg[0]                              1
      '  00203 29                  push next list element from (FuncArg[0])     1
      '  00204 5D    0             FuncArg[0] = next list element from (FuncArg[0
         Debug.Assert .int8 = &H5C:   Debug.Assert .int16 = dummy0
         Debug.Assert .int8 = &H29
         Debug.Assert .int8 = &H5D:   Dim dummy2: dummy2 = .int16

   End If
   
   '  Foreach body (Display N)
   '  00207 3     3             Push value of [N]                            1
   '  00210 35    1 2 3         PRINT 1 Params are above...                  1
      Dim ForEachBody
      ForEachBody = DoProgn(EndPos, level - hasRetValue)
      
'      If hasRetValue Then FasStack.popIntoVoid
      
   'jump back to foreach Condition
   '00215 57    FFFFFFDC      jump [far] over -36 bytes to 184             1

   
   
'     .RestorePos
     
'     .Position = EndPos + 1
   End With
   
 ' Make full
   Dim ForEach_Full
   ForEach_Full = Join(Array("", _
                  GetIndent(-hasRetValue) & ForEach_Expr, _
                  ForEachBody, _
                  GetIndent(-hasRetValue) & TokenClose _
                  ), PartsSeperator)
   
   
   FasFile.outp make_ITEM(ForEach_Full)
   
   
   isForeach = False
   
Exit Sub
DoForEach_err:
Stop
Resume

End Sub


Private Sub DoIf(Last_FasCmdline As FasCommando, FasCmdline As FasCommando, if_Expr)

   ' THEN
      Dim Then_EndPos
      Then_EndPos = FileStream.Position + FasCmdline.Parameters(1)
      
      If FasFile.FasFile_Version < 4 Then Dec Then_EndPos, 2
      
      Dim hasRetValue As Boolean
      hasRetValue = Not (isIf_noRet)
      
      'Debug.Assert hasRetValue = True
      Dim StackBeforeCall
      StackBeforeCall = FasStack.esp
      
      Dim ThenPart
      ThenPart = DoProgn(Then_EndPos, level - hasRetValue, -hasRetValue)
      
      Dim Stack_Delta
      Stack_Delta = FasStack.esp - StackBeforeCall
      Select Case Stack_Delta
         Case 0
            hasRetValue = False
         Case 1
            hasRetValue = True
         Case Else
            Stop ' unknow stackstate
      End Select
      
    ' some how hasRetValue heuristic has failed
'      Debug.Assert hasRetValue = Not (isIf_noRet)

      
      
      If hasRetValue Then FasStack.popIntoVoid
      
   Dim IfFullCommand
      
   'ELSE
      Dim Else_EndPos
      Else_EndPos = Goto_0x57_forward(False, FromEnd:=True)
      
      Dim hasElse As Boolean
      hasElse = Else_EndPos > (FileStream.Position - hasRetValue)
      If hasElse Then
   
       ' Make if Statement
         Dim ElsePart
         ElsePart = DoProgn(Else_EndPos, level - hasRetValue, -hasRetValue)
         If hasRetValue Then FasStack.popIntoVoid
         
         IfFullCommand = Array("", _
               GetIndent(level - hasRetValue) & if_Expr, _
               ThenPart, _
               ElsePart _
               )

         
      Else
         If hasRetValue Then
            Debug.Assert FileStream.int8 = &H1 'Next command should be 'Push NIL'
            FasStack.push make_ITEM("")
         End If
         
         If hasRetValue Then FasStack.popIntoVoid

         IfFullCommand = Array("", _
               GetIndent(level - hasRetValue) & if_Expr, _
               ThenPart _
               )

      End If
      
    ' Finish IF THEN (ELSE)
      IfFullCommand = Join(IfFullCommand, PartsSeperator) _
                     & PartsSeperator & GetIndent(level - hasRetValue) & TokenClose
      Set IfFullCommand = make_ITEM(IfFullCommand)
      
      If hasRetValue Then
         FasStack.push IfFullCommand
      Else
         FasFile.outp IfFullCommand
      End If
      
      isIf = False
End Sub


' Makes a Progn block
' Adds statements(removes them from output)  and last stack item
'
Public Function DoProgn(EndPos, level, Optional ExpectedStackDelta = 0)
'   On Error GoTo DoProgn_err
 
   Dim prognItems
   prognItems = DoBlock(EndPos, level, ExpectedStackDelta)
   
   Dim Part
   ArraySetFirst Part, GetIndent(level + 1) & TokenOpen("progn")
   
   Dim item
   For Each item In prognItems
     ArrayAdd Part, GetIndent(level + 1) & item
   Next
   
   On Error Resume Next
   ArrayAdd Part, GetIndent(level + 3) & FasStack.Current
   'If Err Then Stop
   
   If UBound(Part) > 1 Then
       ArrayAdd Part, GetIndent(level + 2) & TokenClose("progn")
   Else
      ' Just on entry so delete (progn at the beginn
        ArrayRemoveFirst Part
   End If
   
   DoProgn = Join(Part, PartsSeperator)
Exit Function
DoProgn_err:
Stop
Resume
End Function
  
  
Public Function DoBlock(EndPos, level, Optional ExpectedStackDelta = 0)
  
  Dim BeforeProgn
  BeforeProgn = FrmMain.LispFileData.esp
  
  Dim esp_check
  esp_check = FasStack.esp
  
' process  'Then' Branch
  FasFile.InterpretStream_rek FileStream, Target, EndPos, level + 1
 
 'Should be the same before and after
  If ExpectedStackDelta <> (FasStack.esp - esp_check) Then
  
  ' Common uncritical problem is that the HasRetParam for 'Doif' is not correct
  ' (it'll be adjusted right after that call - using the stack delta)
   Debug.Print "Fas_Brancher::DoBlock Expected StackDelta: " & ExpectedStackDelta, _
               " vs " & (FasStack.esp - esp_check), " in  x.." & _
               OffToStr(EndPos)
'               Stop
  End If
 
 
' >>> progn ==
  Dim prognCount
  prognCount = FrmMain.LispFileData.esp - BeforeProgn
  
   
' Get decompilied statements
  DoBlock = FrmMain.LispFileData.popArray(prognCount)
End Function

Public Function IsOr_Check(branchTarget) As Boolean
   With FileStream
   
       .StorePos
        .Position = branchTarget
         
         Dim cmdAtBranchTarget
         cmdAtBranchTarget = .int8
         IsOr_Check = cmdAtBranchTarget = &H6A
      .RestorePos
      
   End With
End Function

Public Sub HandleCond(FasCmdline As FasCommando) ', cond_item_test)
   Brancher_Cond.HandleCond FasCmdline
End Sub

Public Sub HandleAnd(FasCmdline As FasCommando) ', cond_item_test)
   Brancher_Cond.HandleAnd FasCmdline
End Sub

