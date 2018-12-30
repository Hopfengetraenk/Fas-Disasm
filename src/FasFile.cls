VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FasFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const opt_patches As Boolean = 0

Private Const opt_Log_DumpModulVars As Boolean = 0


Private LocalVars
' Fas4 Files are created by
' Vllib.dll/.rsrc/1212/ALCGEN.FSL
' " FAS4-FILE ; Do not change it!"
' "\n;fas4 crunch\n;"


' (vlisp-compile 'st "3darray.lsp")
'st       Standard build mode
'lsm    Optimize and link indirectly
'lsa     Optimize and link directly

'Disable fas encryption
' vlinit.fsl -> Setq *crunch-fasl* nil

Private Const FAS4_FILE_SIGNATURE$ = "FAS4-FILE"
Private Const FAS3_FILE_SIGNATURE$ = "FAS3-FILE"
Private Const FAS2_FILE_SIGNATURE$ = "FAS2-FILE"
Private Const FAS__FILE_SIGNATURE$ = "FAS-FILE"

Private Const LTFAS_FILE_SIGNATURE$ = "AutoCAD LT OEM Product"

Private Const FSL_FILE_SIGNATURE$ = "1Y"

Public FasFile_Version

Private File As New FileStream

Private CreateFilesTemporary As Boolean


' M o d u l  v a r
'The current active stream/module
'currently there only two used
'0 -init/main / FasResourceData .res
'1 -functions / FasFunctionData .fct
Public m_MVar As Variant
Private m_MVars(1) As Variant

'TODO: How to do this:
'Private FasFunctionnamesByIndex
Const Func_TypePrefix$ = "<Func> "


Private FasFunctionnames
Private FasFuncCurrent As FasFunction


Private FasFunctionData As New FileStream
Public FasFunctionDataInterpeted As New Collection

Private FasResourceData As New FileStream
Public FasResourceDataInterpeted As New Collection

Private FasKey 'As New FileStream

Private CodeStart&
Public Offset_CodeStart&
Public Offset_DataStart&

Private KeyStart&
Private IsDecryptingDone As Boolean
Private void As Variant
Private mvarFileName As String 'As New ClsFilename
Private FunctionStreamLength&, FunctionStreamVars&
Private ResourceStreamLength&, ResourceStreamVars&

Public ResourceStream_DefunMain&


Private Const WS_CONTROLCHAR = 0      'Chr( 0..8,b,e,f)...
Private Const WS_WHITESPACE = 1 '     Tab,LF,NewLine,CR,1A,Space
Private Const WS_BLACKSLASH = 4 '      \
Private Const WS_PIPE = 5 '            |
Private Const WS_ALPHANUMMERIC = &HA ' $0123456789ABCD…
Private Const WS_LIMITER = &HB '       !"'(),;?´{}~
Private Const WS_DASH = &HF '          #
Private WHITESPACETABLE


Dim Last_FasCmdline As FasCommando



Public FasStack As New Stack



Public Event initBegin()
Public Event InitDone()
Public Event DecryptingBegin(BytesToProgress&)
Public Event DecryptingProgress(BytesProgressed&, CharDecrypted&)
Public Event DecryptingDone()
Public Event InterpretingBegin(BytesToProgress&)
Public Event InterpretingDone()


Public Event InterpretingProgress(FasCmdlineObj As FasCommando)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long

Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

Private Const RT_RCDATA = 10&

Private Const RT_STRING = 6&
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long



Dim FasCmdline As FasCommando



'Dim Label_Push_col As New Collection
Dim LevelStack As New Stack
Dim level As Integer


Public Property Get Cond_Disable() As Boolean
   Cond_Disable = FrmMain.Chk_Brancher.value = vbChecked
End Property

'Private Property Let Filename(ByVal vData As String)
'    mvarFilename = vData
'End Property

Public Property Get Filename() As String
    Filename = mvarFileName
End Property




'Build a collection called FasResourceDataInterpeted which divides the Fas-Function-Stream into commandlines
'A commandline consists of Command(=Bytecode),Parameters,Disassembled

'With the FasResourceDataInterpeted-object we can later create a raw-Assembly, add, modify functions or write a new fas-datastream
Public Sub InterpretFile()
      
'Stop
'On Error GoTo InterpretFile_err

'Close #1
'Open Me.Filename & ".txt" For Output As 1

'Close #2
'Open Me.Filename & "_.lsp" For Output As 2


Dim Header As New FasCommando

'Make Header
Dim Log_Header As Log_OutputLine
   With Log_Header
      .offset = "offs"
      .Command_Byte = "cmd"
      .Params_Bytes = "p1 p2 p3 p4"
      .DisASM = "Disasm"
      .Description = "Description"
      .Stack = "Stack"
      .DeCompiled = "Decompiled"
   End With

DoLog_OutputLine Log_Header, 0
FileLog_Add String(110, "-")


FrmMain.LV_Log.ListItems.Clear

Set FasFunctionnames = New Collection

Debug.Print vbCrLf & vbCrLf & vbCrLf & time & " " & String(60, "=")
Debug.Print "Interpreting FasResourceData..."
   
      ReDim m_MVar(ResourceStreamVars - 1)
      m_MVars(0) = m_MVar ' Note: regarding command 0x43(init) that's nil => 0
      
      ReDim m_MVar(FunctionStreamVars - 1)
      m_MVars(1) = m_MVar ' Note: regarding command 0x43(init) that's 0 => 1
      
      
      m_MVar = 0
      
'      Dim ResourceStream As New StringReader
      FasResourceData.Position = 0
'      ResourceStream.Data = FasResourceData.Data
'      FasResourceData.CloseFile
      
'      InterpretStream ResourceStream, FasResourceDataInterpeted
      InterpretStream FasResourceData, FasResourceDataInterpeted



Debug.Print "Finish"


Fill_Functionlist

Fill_Stringlist


''next stringtable
''Set m_MVar = m_MVars(2)
'
'Debug.Print String(60, "^")
'Debug.Print "Interpreting FasFunctionData..."
'
'      m_MVar = 1
'
''      On Error Resume Next
''
''      Dim FunctionStream As New StringReader
'      FasFunctionData.Position = 0
''
''      FunctionStream.Data = FasFunctionData.Data
''      If Err Then
''       ' in case of low Memory use stream directly
''         Set FunctionStream = FasFunctionData
''      Else
''
''         FasFunctionData.CloseFile
''
''      End If
''
''      On Error GoTo 0
'
'
''      FasStack.ESP = 0
''      InterpretStream FunctionStream, FasFunctionDataInterpeted
'       InterpretStream FasFunctionData, FasFunctionDataInterpeted
'
'
'
'Debug.Print "Finish"
'
'
'   Fill_Functionlist
'
'   Fill_Stringlist
   
   
   
   SaveDecompiled
   
  'clear storage
   Set FrmMain.LispFileData = New Stack
   
   Err.Clear

'InterpretFile_err:
'Close #1 ' Close log file
'Err.Raise Err, , Err.Description
'Stop
'



End Sub

Private Sub Fill_Functionlist()

   On Error GoTo Err_Fill_Functionlist:
  'Fill Functionlistwindow
   frmFunction.Clear
   Dim objtmp As FasFunction
   
   With frmFunction.Lb_Functions
      For Each objtmp In FasFunctionnames

         .AddItem objtmp.FuncName & "  " & OffToStr(objtmp.Startoffset), 0
      Next
   End With

   frmFunction.Visible = True

Exit Sub
Err_Fill_Functionlist:
   
   frmFunction.Lb_Functions.AddItem "FasFunctionnames ->No functions set - Err: " & Err.Description
   
   frmFunction.Visible = True

End Sub

Private Sub Fill_Stringlist()
   On Error Resume Next
   'Fill Stringlistwindow
   frmStrings.Clear
   
   Dim LV As ListView, li As ListItem, si As ListSubItem
   Set LV = frmStrings.Lv_Strings
   
   With LV
      Dim Counter&
      Dim objtmp1
      Dim Modul
      For Modul = 0 To 1
         Counter = 0
         For Each objtmp1 In m_MVars(Modul)
           ' frmStrings.Lb_Strings.AddItem "#" & BlockAlign_l(Counter, 5) & " " & _
                                             objtmp1 & "  0" & vbTab & TypeName(objtmp1)
            Set li = .ListItems.add(, , "" & Counter)
            
            With frmStrings.LV
               Set si = .ListSubItem(li, "mod")
               si.Text = Modul
               
               Dim TypeColor
               TypeColor = GetColor_Type(TypeName(objtmp1))
               
               Set si = .ListSubItem(li, "val")
                   si.ForeColor = TypeColor
                   si.Text = objtmp1
               
               Set si = .ListSubItem(li, "type")
                   si.ForeColor = TypeColor
                   si.Text = TypeName(objtmp1)
               
            End With
   
            Inc Counter
         Next
         
'      .AddItem String(50, "-")
      
Next

'      For Each objtmp1 In m_MVars(1)
'         .AddItem "#" & BlockAlign_l(Counter, 5) & " " & _
'                                          objtmp1 & "  1" & vbTab & TypeName(objtmp1)
'         Inc Counter
'      Next
      
   End With
   
   frmStrings.Visible = True

End Sub





Private Sub InterpretStream(FileStream As FileStream, Target As Collection)
 ' + FasFunctionData because they are processed together
   RaiseEvent InterpretingBegin(FileStream.Length + FasFunctionData.Length)
   
      InterpretStream_rek FileStream, Target
   RaiseEvent InterpretingDone
End Sub




Public Property Get LVars(idx) As Fas_LVar
   
   On Error Resume Next
   

'   LVars = "L" & idx & "_DecompErr"
   Set LVars = FasFuncCurrent.LocalVars(idx)
   If Err Then
      Set LVars = New Fas_LVar
      LVars.value = idx
   End If
   
   If idx = 0 Then Exit Property
'   Debug.Print "Get LVars(" & idx & " ):" & FasFuncCurrent.LocalVars(idx) & " " & FasFuncCurrent.FuncName
'   Debug.Print "     MVars(" & idx & ",0):", MVarsEx(idx, 0)
'   Debug.Print "     MVars(" & idx & ",1):", MVarsEx(idx, 1)
'asd
'Stop
'asd
'qw

End Property

Public Property Let LVars(idx, ByVal NewValue As Variant)
On Error Resume Next
'   Debug.Print "Let LVars(" & idx & " ):" & FasFuncCurrent.LocalVars(idx), "<=", NewValue
'   Debug.Print "     m_MVars(0)(" & idx & "):", m_MVars(0)(idx)
'   Debug.Print "     m_MVars(1)(" & idx & "):", m_MVars(1)(idx)
   
   
   FasFuncCurrent.LocalVars(idx) = NewValue
   If Err Then
      Dim New_LVars As Fas_LVar
      Set New_LVars = New Fas_LVar
      New_LVars.value = idx
      
      Set FasFuncCurrent.LocalVars(idx) = New_LVars
   End If
   
   
'   If idx = 0 Then Exit Property
'Stop

End Property




Public Property Get MVarsEx(idx, table) As Variant
   
   MVarsEx = "MVar#" & idx
   
   Dim tmp
   Set tmp = m_MVars(table)(idx)
'
'   If Not (IsEmpty(tmp)) Then
      Set MVarsEx = tmp
'   End If
   
   
   
'Planned
' to decompile when there is some (setq <Func>)
'   If isFunction(tmp) Then
'      FrmMain.AddtoLog "Function tiggered - Doing Decompile for" & tmp
'      FasCmdline.Interpreted = FasCmdline.Interpreted & ";; direct decompile not implementent   Please manually cut&Paste function from down to here"
'
'     ' Finds FasFunctionbyOffset
'
'     ' TODO: Improve code/organisation
'     ' Cutting out offset again from String we created before is bad (performance / maintainibility).
'       Dim item As FasFunction
'       Dim Startoffset$
'       'tmp example
'       '"<Func> -lambda- Modul:LV_0_DecompErr, Offs: $0 [0]"
'       Startoffset = Split(Split(tmp, "Offs: ")(1))(0)
'       Debug.Assert Left(Startoffset, 1) = "$"
'       Startoffset = CStr(OffToVal(Startoffset))
'
'       Set item = FasFunctionnames(Startoffset)
'
'       With FasFunctionData
'         .StorePos
'
'
'         .RestorePos
''      FasFunctionData
'  '    InterpretStream_rek(FasFunctionData,
'      End With
'   End If

End Property

Public Property Let MVarsEx(idx, table, ByVal NewValue As Variant)
   On Error Resume Next
   Set m_MVars(table)(idx) = NewValue
   If Err Then
      Debug.Print "Let MVarsEx Failed " & Err.Description
      Debug.Print , "Loaded m_MVars(" & table & ")(" & idx & ") with  m_MVars(0) " & UBound(m_MVars(0)) & "  m_MVars(1) " & UBound(m_MVars(1))
   End If
End Property


Public Property Get MVars(idx) As Variant
   On Error Resume Next
   
   Set MVars = MVarsEx(idx, m_MVar)
   
'   MVars = Join(MVarsEx(idx, m_MVar))
If Err Then Stop
End Property

Public Property Let MVars(idx, ByVal NewValue As Variant)

   MVarsEx(m_MVar, idx) = NewValue

End Property


'TODO: Tag object with Type explicitly
' Like for ex of it was created by command '&h39' it'll be a list
Public Function isList(Expression) As Boolean
   If IsObject(Expression) Then
   
      isList = TypeOf Expression Is T_LIST
      
   Else
   
   'crumsome unreliable implicit Type recog.
      Dim tmp
      tmp = LTrim(TokenRemove(Expression))
      
      Dim char
      char = Left(tmp, 1)
      If char = "(" Then
         isList = isList(Mid(tmp, 1))
       Else
         isList = IsNumeric(char)
       End If
       
   End If
End Function


Public Function isFunction(Name) As Boolean
   On Error Resume Next
   '           Const Func_TypePrefix$ = "<Func> "
   isFunction = Func_TypePrefix = Left(Name, Len(Func_TypePrefix))

End Function


Public Sub asValue(item)
' LISTs and SYMs will NOT be shown with a >'< at the beginning
  If TypeOf item Is T_LIST Or _
     TypeOf item Is T_SYM Then
     item.AsSYM = False
  End If
End Sub
Public Sub asSymbol(item)
' LISTs and SYMs will be shown with a >'< at the beginning
  If TypeOf item Is T_LIST Or _
     TypeOf item Is T_SYM Then
     item.AsSYM = True
  End If
End Sub

' Output item
Public Sub outp(item)
   With FasStack
   
    ' Enqueue item for output
    ' -> Push item on stack
    ' -> Ouput it directly if stack is empty
      If .esp = 0 Then
          FasCmdline.Interpreted = item
          RaiseEvent InterpretingProgress(FasCmdline)
      Else
          .Current = item
      
      End If
   End With
End Sub
Public Function IsNextCommand(FileStream As FileStream, ExpectedValue As Byte, Optional seekback = True) As Boolean
   With FileStream
      Dim NextCmd
      NextCmd = .int8: If seekback Then .Move -1
      IsNextCommand = ExpectedValue = NextCmd
   End With
End Function


Public Sub InterpretStream_output( _
                                    DisASM As String, _
                                    DisASM_short, _
                                    Description _
                                 )
'     FasCmdline.Interpreted = Interp
      FasCmdline.Disassembled_Short = DisASM_short
      FasCmdline.Disassembled = GetIndent(level) & DisASM
      FasCmdline.Interpreted = GetIndent(level) & FasCmdline.Interpreted
      
      RaiseEvent InterpretingProgress(FasCmdline)
      
End Sub
Private Function CollectionQuery(Collection As Collection, Key$)
   ' Create and Get itemCounter
     If Not Collection_IsAlreadyIn(Collection, Key) Then
         Collection.add make_ITEM(0), Key
     End If
     Set CollectionQuery = Collection(Key)
End Function

Public Sub InterpretStream_rek(FileStream As FileStream, Target As Collection, Optional StopAtOffset = &H7FFFFFFF, Optional level = 0, Optional Stackadd = 0, Optional ANDDest = 0, Optional InterpretStream_AND_6A_Helper = 0)

On Error GoTo InterpretStream_rek_err
   Dim tmp, tmp2, tmp3, dummy, ax As Long, cx As Long, ebx, edi
   Dim DisASM As String, DisASM_short As String
   Dim Description As String

   Dim FunctionToDecompile As FasFunction
   Dim LambdaOpenParatheseCount&

   
   
   
   Dim StartEsp&
   StartEsp = FasStack.esp
   FasStack.esp = FasStack.esp + Stackadd
   
 ' Create brancher obj
   Dim Brancher As New Fas_Brancher
   Set Brancher = New Fas_Brancher
   Brancher.create FileStream, FasStack, Me, level
   
   With FileStream
      
      
      Do Until .EOS Or (.Position >= StopAtOffset)
         
       ' Create FasCmdline Object
         Set FasCmdline = New FasCommando
            
            
         Set FasCmdline.FasFile = Me
         
       ' Store current Offset
         FasCmdline.Position = .Position
         
       ' Initialise FasCmdline Object
         FasCmdline.ModulId = m_MVar

'         Debug.Assert FasCmdline.Position < &H81
       
         FasCmdline.Commando = .int8
'         If (FasCmdline.Position = 0) And ((FasCmdline.Commando & &HFD) <> &O14) Then
'            MsgBox "it look like there went something wrong with opening or decrypting this file"
'         End If

         
'         If Not Last_FasCmdline Is Nothing Then
'         ' next 2 lines Not working on primitives(variant)
'            Set FasCmdline.Stack_Before = Last_FasCmdline.Stack_After
'            Set FasCmdline.Stack_Pointer_before = Last_FasCmdline.Stack_Pointer_After
'         End If
         
         Set FasCmdline.Parameters = New Collection
         Dim Param As Collection
         Set Param = FasCmdline.Parameters
         
       ' add FasCmdline Object to Target collection
'         Target.Add FasCmdline
         FasCmdline.Interpreted = ""
       
         DisASM_short = ""
         Description = ""
       
'         Label_Pop FasCmdline.Position
       
       ' Process Commands
         Select Case FasCmdline.Commando
         
'=== Defun ===  FunctionHead
         'defun 'defun-q
         Case &H14, &H15:
            DisASM_short = "DEFUN"
            Description = "Subroutine; Setup frame for locals ;4xI8 "
            
            'Check for bad start (previous commands)
            'Debug.Assert FasStack.ESP <= 2
            'FasStack.ESP = 0
            
            Dim isFas As Boolean
            isFas = FasCmdline.Commando = &H14

            
            
          ' --- Get Params ---
            
            On Error GoTo InterpretStream_rek_err
            
          ' GetParams (4x Byte)
            Dim p1_LVrs, p2_Args, p2_Args_Max, p2_GC
            Param.add make_INT(.int8, 1)
            
               Param.add make_INT(.int8, 1)
               p2_Args = Param(2)

               Param.add make_INT(.int8, 1)
               p2_Args_Max = Param(3)

               

            Param.add make_INT(.int8, 1) 'and         eax,0FFFE
            ax = Param(4)
            cx = Param(1)
          ' ax and cx represents a 15 bit value with the order axcx
          ' the last bit of ax is discarded
            p1_LVrs = ((ax And &HFE) * &H80) Or cx
'            Debug.Assert p1_LVrs <> 0
            
          
          ' only last bit matters
            p2_GC = CBool(ax And &H1)
            
            
            
          ' --- GetName ---
            
           'Preinit Name heuristically
            If m_MVar = 0 Then
               Dim FasFunction As FasFunction
               Set FasFunction = New FasFunction
               
               FasFunction.Startoffset = FasCmdline.Position
               
               FasFunction.ModulId = m_MVar 'Res
               Set FasFunction.ModulStream = FasResourceData
               
               Set FasFuncCurrent = FasFunction
   
             ' TODO: Lookup the real names of these two function in the FSL-Files
               If p2_Args = 1 Then
               
               '# PreInit for ThisStream::Loader
                  Set FasFuncCurrent = New FasFunction
                  FasFuncCurrent.FuncName = "fs_init"
                  
                  Dim Arg1$
                  Arg1 = "funcStream"
                  FasFuncCurrent.Args.add Arg1
                  
                  LVars(0) = Arg1
                  
               Else
               
               '# PreInit for ThisStream::Main
                  Set FasFuncCurrent = New FasFunction
                  FasFuncCurrent.FuncName = "main"

                  
                ' Memorize Start of Main to seek there after decompiling
                  ResourceStream_DefunMain = FasCmdline.Position
               End If
                  
            Else
              'to indicate something when wrong with lookup
               DisASM = "ERR_KNOWN_FUNC_NAME"
            End If
            
            Dim bSuppressDecompOutput As Boolean
            bSuppressDecompOutput = m_MVar = 0
            
          '  ...using the >offset< to loop it up in  => FasFunctionnames[]
            On Error Resume Next
            Const ERR_InvProcCall& = 13
            Const ERR_ObjVarOrWithBlockNotSet& = 91
            
          ' Look Functionname
            Dim FuncName
            Set FasFuncCurrent = FasFunctionnames(CStr(FasCmdline.Position))
            FuncName = FasFuncCurrent.FuncName
            
            If (bSuppressDecompOutput = False) And (Err = ERR_InvProcCall Or Err = ERR_ObjVarOrWithBlockNotSet) Then
               FrmMain.AddtoLog "Warning: Couldn't get func name for (defun @" & OffToStr(FasCmdline.Position)
            End If
            
          ' Keyword
            Dim Keyword
            Keyword = IIf(isFas, "defun", "defun-q")
            
            If (Not FasFuncCurrent Is Nothing) Then
               If FasFuncCurrent.isLambda Then
                  Keyword = "lambda"
                  FuncName = ""
               End If
            End If

          ' Process Arguments
            Dim Args As clsStrCat: Set Args = New clsStrCat
            If p2_Args = 0 Then
             ' to show that there are no more args
               Args.Concat TokenFull("")
            Else
            
               Args.Concat TokenOpen("")
               
             ' Concat args
               With FasFuncCurrent: Dim arg
               If .Args.count Then
               For Each arg In .Args
                        Args.Concat arg & " "
               Next: End If: End With
               
               Args.Concat TokenClose("", level)

            End If
            
          ' Make Function definition
            DisASM = TokenOpen(Keyword) & FuncName & " " & Args.value
            
            
          ' Chanced for optimising function + param output
            If bSuppressDecompOutput = False Then FasCmdline.Interpreted = DisASM
'            FasStack.push "(" & disasm

          ' --- Fill details about Function
            FasFuncCurrent.NumOfArgs = p2_Args
            FasFuncCurrent.VarsCount = p1_LVrs
          ' --- CreateLocals ---
                     
            FasFuncCurrent.MakeLVars p1_LVrs, IIf(isFas, "_FAS", "_FSL")
         
            
          ' --- Output ---
            DisASM = Join(Array(DisASM, "  ", _
                     "LVrs:", p1_LVrs, " ", _
                     "Args:", p2_Args & ".." & p2_Args_Max, " ", _
                     "GC:", boolToText(p2_GC) _
                     ))
                     
         
         Case &H17
            'seems to be used in FAS2-File only
            
            DisASM = "main"
            On Error Resume Next
            DisASM = FasFunctionnames(CStr(.Position - 1))
            On Error GoTo InterpretStream_rek_err
            
            DisASM = TokenOpen("defun") & DisASM & " "
            
            FasCmdline.Interpreted = DisASM
            
            DisASM = "FAS2 & disasm "
            
            Param.add make_INT(.int16, 2) '
            DisASM = DisASM & MVars(Param(1))
            'push clng(3)
         
         
'Const nil & T
         Case &H1
            DisASM_short = "ld NIL"
            Description = "Pushs NIL on the Stack"
            
            DisASM = "Push nil"
            Set tmp = make_NIL()
            
            FasStack.push tmp ' or '()
            
         Case &H2
            FasStack.push make_ITEM("T")
            DisASM = "Push T" 'command AND?


' === Local Vars ===
            
        'Get _FSL
'         Case &H5: 'push Local Var8 on Stack '00000101
'            param.add make_INT(.int8, 1)           ' GetVarIdx
'            tmp = LVars(param(1))     ' GetValue
'            DisASM = "push_FSL LVar" & tmp & "  push Local Var8 on Stack"
'            FasStack.push tmp           ' Push Value onto Stack

        'Set _FSL
         Case &H8: 'FSL pop into Local Var8 '00001000
            Param.add make_INT(.int8, 1)         ' GetVarIdx
            tmp = FasStack.pop      ' GetValue from Stack
            DisASM = "LVar8_" & LVars(Param(1)) & " = " & tmp
            LVars(Param(1)) = tmp   ' Assign Value
            
            'FasCmdline.Interpreted = disasm
            
        'Zero _FSL
         Case &H64: 'FSL clear Local Var8 '01100100
            Param.add make_INT(.int8, 1)         ' GetVarIdx
            DisASM = "clear LVar8_" & LVars(Param(1))
            LVars(Param(1)) = "<Deleted LVar8>"    'Set Value to Zero
         


        'LVar16=[ebp-18h] or [local.6] <= LVar
         
        'Get _FAS
         Case &H5, &H5C 'push Local Var16 on Stack '01011100
            Dim bisFas As Boolean
            bisFas = FasCmdline.Commando = &H5C

            DisASM_short = IIf(bisFas, "F_", "") & "getVAR"
            Description = "Push global Var => stack; 1x Param"
            
            Dim LVarIdx
            LVarIdx = .int8 ' GetVarIdx
            
          ' in Fas-Version this value is 16 bit
            If bisFas Then
            
               Debug.Assert .int8 = 0 ' Get highPart of VarIdx
               'If not null get as 16 bitvalue

            End If

            Param.add LVarIdx
            
            Dim LVarValue
            Set LVarValue = LVars(LVarIdx)       ' GetValue
            
            DisASM = "L_" & LVarIdx & " => " & LVarValue

            FasStack.push LVarValue           ' Push Value onto Stack
         'don't forget to adapt 'case 43' also if you modify this
         
        'Set _FAS
         Case &H5D 'pop into Local Var16 '01011101
          ' GetVarIdx
            DisASM_short = "setVAR"
            Description = "Pop global stack => Var; 1x Param"

            Param.add make_INT(.int16, 2)         ' GetVarIdx
            LVarIdx = Param(1)
            
            LVarValue = FasStack.pop       ' GetValue from Stack
            
            DisASM = "L_" & LVarIdx & " <= " & LVarValue
            LVars(LVarIdx) = LVarValue     ' Assign Value

        'Zero _FAS
         Case &H5E 'clear Local Var16
            Param.add make_INT(.int16, 2)         ' GetVarIdx
            DisASM = "clear L_" & LVars(Param(1))
            LVars(Param(1)) = "<Deleted LVar>"   'Set Value to Zero
         

'=== BinValues ops ===

         'Get
         Case &H3  'Push GVvar16 on stack
           DisASM_short = "VALUE"
           Param.add make_INT(.int16, 2)  'String?
           tmp = MVars(Param(1))
           asValue tmp
           FasStack.push make_ITEM(tmp)
           
           DisASM = "Push value of [" & tmp & "]"

           

         Case &H4  '         'FSL
            'goes together with 2e
             Param.add make_INT(.int8, 1)
             Param.add make_INT(.int8, 1)
             tmp = "stream " & Param(2) & " [" & Param(1) & "]"
             DisASM = "push " & tmp 'lnode*
             'Debug.Assert param(1) = 1 'listselector or rewinder - seems to be always 1
             'Debug.Assert (param(2) = 0) Or (param(2) = 1) Or (param(2) = 2) Or (param(2) = 4)
             
             FasStack.push tmp
             
             
         'setq
         Case &H6, &H1A, &H1B: ' Pop into GVar16
         
            If FasCmdline.Commando = &H6 Then
               DisASM_short = "setq"
               
            ElseIf FasCmdline.Commando = &H1A Then
               DisASM_short = "FSL_defun"             'setq_FSL_func
               
            ElseIf FasCmdline.Commando = &H1B Then
               DisASM_short = "FSL_setq"              'setq_FSL_Var
               
            End If
            
            Description = "Push assignment: param#1 = last stack item"

            Param.add make_INT(.int16, 2)
            
            Dim AssignToSymbol
            Set AssignToSymbol = MVars(Param(1))
            
            asValue AssignToSymbol
            
            Dim Expression
            Set Expression = FasStack.Current
            FasStack.popIntoVoid
            
            
            
            Debug.Assert TypeOf AssignToSymbol Is T_SYM
            '...with exploring intent. Just  for finding special cases (or errors)
            
            
               '-- Make setq --
   '            On Error Resume Next
               DisASM = AssignToSymbol & " = " & Expression.toText
   '            DisASM = Join(tmp) & " = " & DisASM
   
   
       ' #Patch: Setq_PUSH-POP-supress  'Echo' after setq
         If opt_patches Then
             ' consume next  pu_item <FuncName>
               Dim pu_item, pu_item_Idx
               pu_item = .int8
               pu_item_Idx = .int16
               
               If (pu_item <> 9) Or (pu_item_Idx <> Param(1)) Then
               ' Normal defun is done with h09_puitem <idx>; h06_Setq <idx>
               ' ErrorHandler is done with h03_puitem <idx>; h06_Setq <idx>
               '  Stop
                 .Move -1 - 2
               Else
                  FasStack.push "<dummy>"
               End If
                 
            End If
   
          ' defun-Handler#2b: for possible defun
            If Not FunctionToDecompile Is Nothing Then
               FunctionToDecompile.isLambda = False 'Not really need since 'false' is the default value
               
             ' Just checking - if the name of function match
               Debug.Assert FunctionToDecompile.FuncName = AssignToSymbol
               
             ' That is no real setq - suppress output to make space for the defun
             ' that will follow
               tmp3 = ";;; (" & AssignToSymbol & " @" & OffToStr(FunctionToDecompile.Startoffset)
               
               
            Else
               tmp3 = TokenFull(DisASM_short, AssignToSymbol, Expression.toText)
            End If
   
   
               Set tmp = New E_SETQ
               tmp.value = tmp3
               
               If TypeOf Expression Is E_ITEM Then
                  If Expression.NoOutput Then
                     tmp = ""
                     DisASM = ";;; " & DisASM
                  End If
               End If

               
   '            On Error GoTo InterpretStream_rek_err
               
               'To decide whether the line is complete and should be outputed or to go on concatinating
               'StartEsp is used since this function can be called by it self again
               'but uses global stack that doesn't starts at StartEsp
               If FasStack.esp > (StartEsp) Then
                  FasStack.Current = tmp
               Else
                  FasCmdline.Interpreted = tmp
               End If
               
'            End If
            
            
'         'setq_FSL_func"
'         Case &H1A ' Pop into GVar16
'            DisASM_short = "FSL_defun"
'            param.add make_INT(.int16, 2)
'            DisASM = MVars(param(1)) & " " & FasStack.pop
'
'            FasCmdline.Interpreted = TokenFull("FSL_defun", DisASM)
'
'
'         'setq_FSL_Var
'         Case &H1B ' Pop into GVar16
'            DisASM_short = "FSL_setq"
'            param.add make_INT(.int16, 2)
'            DisASM = MVars(param(1)) & " " & FasStack.pop
'
'            FasCmdline.Interpreted = TokenFull("FSL_setq", DisASM)
'
            
        'Copy Element
         Case &H7:
            Param.add make_INT(.int8, 1) 'List(>index<)
            Param.add make_INT(.int8, 1) 'List(>index<)
            DisASM = "Copy Element " & Param(1) & " to " & Param(2) & " ListBase=[61A49790]"
            
            'Djamana: Attention I just added this to make the stack more nice
            '         but there maybe no real coderelated reason for that - so maybe remove it.
            FasCmdline.Interpreted = FasStack.pop
            '61A49790
  
  
         Case &H10 'List Step
            Param.add make_INT(.int8, 1) 'steps forward or backward in a list ?
            Param.add make_INT(.int8, 1)
            DisASM = "Push??? LVar " & Param(1) & ", " & Param(2)
            FasStack.push DisASM
  
         
         Case &H9: 'Push GVar16 on Stack
            DisASM_short = "pu_Item"
            
            'Used for Lists and Strings
            Param.add make_INT(.int16, 2) '.int16Sig
            
            
            On Error Resume Next
            Set tmp = MVars(Param(1))
            If Err Then
               Err.Clear
               tmp = MVars(Param(1))

            Else


               asSymbol tmp

            End If
            'tmp.value = tmp.toText
            
            On Error GoTo InterpretStream_rek_err:
            
            DisASM = "push " & tmp
            Description = "push item #" & Param(1) & " from GVar => Stack"
         
          ' defun-Handler#1:  Memorise function to decompile for next stack output
            Dim isDefun As Boolean
            isDefun = TypeOf tmp Is T_USUBR
            If isDefun Then
               
               Set FunctionToDecompile = FasFunctionnames(CStr(tmp.Start))
               LambdaOpenParatheseCount = 0
               
             ' Clear FunctionText
             ' ^- maybe move this to 'defun-Handler#2a:
               Set tmp = make_ITEM("")
               
            
            End If
            
            
            FasStack.push tmp
            
            
            
         Case &HA 'pop dummy
            DisASM_short = "Pop"
            Description = "Pop/Remove last element from stack"
            
            DisASM = "pop dummy (decrease stack)" & vbCrLf  ' "sub-4"
            FasCmdline.Interpreted = FasStack.pop
            
'            Debug.Assert FasStack.esp = 0 ' at this point the stack should be always embalanced
                                          ' else this indicated some problem
            
          ' Error Cleanup
            Do Until FasStack.esp = 0
               FasCmdline.Interpreted = FasCmdline.Interpreted & _
                     ";;;<Error app='FAS-DISAM' hint='GarbageFromStack'>" & FasStack.pop & "</Error>"
            Loop
            
            
            
         Case &HB: 'Stack Duplicate Element
            DisASM_short = "Pu_Last"
            Description = "Push last element again on stack"
            
            FasStack.push make_ITEM(FasStack.Current)
            DisASM = "Push " & FasStack.Current & "[Last element again]"
              
         Case &HC 'Push Gvar16 onStack
            Param.add make_INT(.int16, 2) '.int16Sig
            FasStack.push MVars(Param(1))
            DisASM = "FSL_Push " & FasStack.Current
            
            
  '==== Branches ===
            
         Case &HD, &H3C ' Branch_16 if false
         ' identical code in vl.arx (not point same points in cmd-switch table)
         '     0d -> 0f_jmp only used in fsl and fas2
         '     0f -> 3C_with_jump_backward
         '0d - Somehow use to jump over errorhandler

            DisASM_short = "BrIfF16"

          ' param #1 - branchDelta16
            Dim branchDelta
            branchDelta = .int16
            Param.add make_INT(branchDelta, 2)

          ' param #2 - Condition (from stack)
            Dim Condition
            Condition = FasStack.pop

            Dim branchTarget&
            branchTarget = .Position + branchDelta

            DisASM = "if (" & Condition & "==0) then jmp to " & OffToStr(branchTarget)
            FasStack.push "'Retval' of IF"

            InterpretStream_output DisASM, DisASM_short, Description
            

            If branchDelta > 0 Then
                           dummy = FasStack.pop
               tmp = Brancher.DoProgn(branchTarget, level + 1, 0)

               FasCmdline.Interpreted = TokenFull("Cond? " & Condition, tmp)
               'InterpretStream_rek FileStream, Target, branchTarget, level + 1, -1
'               If .int8 = &H1A Then 'is next cmd setq?
'                  qw
'                  param.add make_INT(.int16, 2)
'                  disasm = "FSL_setq " & MVars(param(1)) & " " & FasStack.pop
'
'               Else
'                  .Move -1
                 'FasStack.pop
'                 GoTo NoOutput
'               End If

            Else
               FasStack.popIntoVoid
               GoTo NoOutput
            End If

         
         Case &HE, &H3D ' Branch_16 if true
         ' identical code in vl.arx (not point same points in cmd-switch table)
         
                   ' param #1 - branchDelta16
            'Dim branchDelta
            branchDelta = .int16
            Param.add make_INT(branchDelta, 2)

          ' param #2 - Condition (from stack)
            'Dim Condition
            Condition = FasStack.Current
            
            'Dim branchTarget&
            branchTarget = .Position + branchDelta
            
            '3d - seen in fsl
            DisASM_short = "BrIfT16" & IIf(FasCmdline.Commando = &HE, "_2", "")
            
            DisASM = "if (" & Condition & ") then pop else jmp to " & OffToStr(branchTarget)
            
            InterpretStream_output DisASM, DisASM_short, Description

            tmp = Brancher.DoProgn(branchTarget, level + 1, 1)
            'InterpretStream_rek FileStream, Target, branchTarget, level + 1, -1
            
            FasStack.popIntoVoid
            FasCmdline.Interpreted = TokenFull("Cond2? " & Condition, tmp)
            
        '    If level = 0 Then FasStack.push "'Retval' of IF"
        
'            qw
'            GoTo NoOutput
         
         
         Case &H67: 'Branch_32 if true
         ' identical code in vl.arx (not point same points in cmd-switch table)
         '     0d -> 0f_jmp only used in fsl and fas2
         '     0f -> 3C_with_jump_backward
                   
           DisASM_short = "BrIfT"
         'normal if then else 57 used for
         '( while testexpr [expr])
         '( repeat int [expr])
         '( foreach name lst [expr]) and
         '( if testexpr thenexpr elseexpr)
            
            
            Dim isBrIfT_67 As Boolean
            isBrIfT_67 = FasCmdline.Commando = &H67
            
            If isBrIfT_67 Then
            
             ' param #1 - Jump32
               branchDelta = make_INT(.int32Sig, 4)
               DisASM_short = DisASM_short & "32"
            Else
               branchDelta = make_INT(.int16, 2)
               DisASM_short = DisASM_short & "16"
            End If
            
            Param.add branchDelta
            
            
            'Dim Condition
            Condition = FasStack.pop
            
            'Dim branchTarget&
            branchTarget = .Position + branchDelta
            
               'Dim Keyword
               Keyword = Brancher.GetKeyWord(Last_FasCmdline, FasCmdline)
               
              'VlaxFor Patch
               If Brancher.isVlaxFor Then Condition = Brancher.VlaxFor_item
               
            
               DisASM = Keyword & " (" & Condition & ") pop else goto " & OffToStr(branchTarget)  ' If [edi]==0
            
            
            'FasCmdline.Interpreted = condition
            'FasStack.pop
            'FasStack.push condition
            
            Dim Branch_Firstpart
            FasCmdline.Disassembled = GetIndent(level) & DisASM
            Branch_Firstpart = TokenOpen(Keyword, level) & Condition
            
            'If Keyword <> "if" Then
            '   FasCmdline.Interpreted = Branch_Firstpart
            'Else
            '   'FasStack.push Branch_Firstpart
            'End If

            InterpretStream_output DisASM, DisASM_short, Description

            
           ' If Brancher.isRepeat Then
           '    .Move 1 '<-for repeat ; Skip some 0x1 Push nil
           ' End If
            
            Brancher.Decompile Last_FasCmdline, FasCmdline, Branch_Firstpart
            


            
 
'          ' is it normal_If?
'            If Last_FasCmdline.Commando = &H57 Then 'Note: &H57 is Goto xxx
'             ' 'else' branch of if
'               If Last_FasCmdline.Parameters(1) >= &H0 Then
'
'                  InterpretStream_rek FileStream, Target, Last_FasCmdline.Position + Last_FasCmdline.Parameters(1), level + 1
'                  'FasStack.pop
'                  'FasStack.Current = "Then OR Else"
'               Else
'                ' is it While  (For Each)?
'                  If bRepeat Then
'                   '  FasStack.Current = "Repeat"
'                  Else
'                     Dim NextCmd&
'                     NextCmd = .int8
'                     .Move -1
'                     If NextCmd = &H19 Then 'Note: 19    clear args+vars
'                    '    FasStack.Current = "For Each"
'                     Else
'                     '   FasStack.Current = "While"
'                     End If
'
'                  End If
'               End If
'            Else
'               Err.Raise vbObjectError, "", "Then-block of normal_if(&h67) did end with goto(&H57)"
'               Stop
'            End If
            
        
'            FasCmdline.Interpreted = FasStack.pop & ")"
 '           disasm = GetIndent(level)  & "END IF"
            GoTo NoOutput
            
         Case &H69 'Branch32
            DisASM_short = "Br2IfT32"
            
          ' param #1 - Jump32
          ' Dim branchDelta
            branchDelta = make_INT(.int32Sig, 4)
            Param.add branchDelta
               
           'Dim Condition
            Condition = FasStack.pop
          
           'Dim branchTarget&
            branchTarget = .Position + branchDelta
               
               
            Param.add make_INT(.int32Sig, 4)
            'pop condition always
            DisASM = "If (" & Condition & ") ... else goto " & OffToStr(branchTarget)
         
            'seldom used - find out more
            Stop
         
         
         Case &H68, &H6A 'Cond
         
            Dim isAND_6A As Boolean
            isAND_6A = FasCmdline.Commando = &H6A
         
          ' param #1 - branchDelta32
            branchDelta = .int32Sig
            
          ' param #2 - Condition (from stack)
            Param.add make_INT(branchDelta, 4)
            branchTarget = .Position + Param(1)
            
            
            If isAND_6A Then
               DisASM_short = "AND"
               DisASM = "and_If (" & FasStack.Current & ") .pop Else goto " & OffToStr(branchTarget)

            Else
               DisASM_short = "CND/OR"
               DisASM = "cond/or If (" & FasStack.Current & ") Goto " & OffToStr(branchTarget) & " Else .pop"
            End If
            
'            Label_Push CStr(branchTarget), DisASM_short
               InterpretStream_output DisASM, DisASM_short, Description
             
            If isAND_6A Then
               '=== AND ===
               '   $14E   3  12   4 VALUE     Push value of [T]
               '   $151  6A  17   4 AND       and_If (T) .pop Else goto $167
               '
               '   $156   3  11   4 VALUE     Push value of [B]
               '   $159  6A  9    4 AND       and_If (B) .pop Else goto $167
               '
               '   $15E   3  14   4 VALUE     Push value of [C]
               '   $161  6A  1    4 AND       and_If (C) .pop Else goto $167
               '
               '   $166   2       4           Push T
               
               
               
               'Brancher.HandleAnd FasCmdline
               ' Create and Get itemCounter
                 Static And_ItemsOnStack As New Collection
                 Dim And_ItemsCount
                 Set And_ItemsCount = CollectionQuery(And_ItemsOnStack, CStr(branchTarget))
                 
                 
               ' Increase Item Counter
                 And_ItemsCount.value = And_ItemsCount + 1
                 
               ' Check for Terminator
                 Dim isLastAnd
                 isLastAnd = IsNextCommand(FileStream, &H2) 'Push T
                 isLastAnd = isLastAnd And (branchDelta = 1)
                 '6A  1    AND       and_If (NIL)
                 '02                  Push T
                 
                 If isLastAnd Then
                  
                     Dim And_Items
                     And_Items = FasStack.popArray(And_ItemsCount)
                     And_ItemsCount.value = 0
                     
                     Dim and_Full
                     and_Full = TokenFull("and", And_Items)
                     FasStack.push make_ITEM(and_Full)
                     
                     'Seek over Terminator
                  '   $166   2       4           Push T
                     
                     tmp = .int8
                     Debug.Assert tmp = 2 ' Push T
                 End If

               
            ElseIf Brancher.IsOr_Check(branchTarget) Then
                ' === OR ===
                
                  ';item #1/3
                  '   $12B   3  15   1 VALUE     Push value of [A]
                  '   $12E  68  15   1 CND/OR    cond/or If (A) Goto $142 Else .pop
                  '
                  ';item #2/3
                  '   $133   1       2 ld NIL    Push nil
                  '   $134  68  9    2 CND/OR    cond/or If (NIL) Goto $142 Else .pop
                  '
                  ';item #3/3 -Last item
                  '   $139   3  14   3 VALUE     Push value of [C]
                  '   $13C  68  1    3 CND/OR    cond/or If (C) Goto $142 Else .pop
                  '
                  ';Terminator
                  '   $141   1       4 ld NIL    Push nil
                  '   $142  6A  1    4 AND       and_If (NIL) .pop Else goto $148
                  '   $147   2       4           Push T
                  '
                  '
                  '   $148  35  1 13 4 ld_USUBR  PRINC 1 Params are above...
                  '   $14D   A       3 Pop       pop dummy (decrease stack)
                
                
               ' Create and Get itemCounter
                 Static Or_ItemsOnStack As New Collection
                 Dim Or_ItemsCount
                 Set Or_ItemsCount = CollectionQuery(Or_ItemsOnStack, CStr(branchTarget))
                 
                 
               ' Increase Item Counter
                 Or_ItemsCount.value = Or_ItemsCount + 1
                 
               ' Check for Terminator
                 Dim isLastOR
                 isLastOR = IsNextCommand(FileStream, &H1) 'Push nil
                 isLastOR = isLastOR And (branchDelta = 1)
                 '6A  1    AND       and_If (NIL)
                 '02                  Push T
                 
                 If isLastOR Then
                  
                     Dim Or_Items
                     Or_Items = FasStack.popArray(Or_ItemsCount)
                     Or_ItemsCount.value = 0
                     
                     Dim Or_Full
                     Or_Full = TokenFull("or", Or_Items)
                     FasStack.push make_ITEM(Or_Full)
                     
                     'Seek over Terminator
                  '   $141   1       4 ld NIL    Push nil
                  '   $142  6A  1    4 AND       and_If (NIL) .pop Else goto $148
                  '   $147   2       4           Push T
                     
                     
                     Debug.Assert .int8 = 1
                     Debug.Assert .int8 = &H6A: Debug.Assert .int32 = &H1
                     Debug.Assert .int8 = 2
                 End If


                  
            Else
               
               InterpretStream_output DisASM, DisASM_short, Description
            
             ' === Cond ===
               Brancher.HandleCond FasCmdline ', Condition

               GoTo NoOutput
            End If
            
            'Cond No Brancher
'
'
'                     Dim isLastCond As Boolean
'                     Dim hasRetVal As Boolean
'                     isLastCond = Brancher.Brancher_Cond.isCond_Last(hasRetVal)
'
'
'                   ' that's important to embalence the stack
'                   ' in case of problems use 'FasStack.Current'
'                   ' That might make the stack getting bigger(in case of cond with retval)
'                     If hasRetVal Then
'                        Condition = FasStack.pop
'                     Else
'                        Condition = FasStack.Current
'                     End If
'
'
'                   ' Check if hasRetVal stays at the same value during one cond
'                     Static hasRetVal_Last
'                     If IsEmpty(hasRetVal_Last) = False Then _
'                        Debug.Assert hasRetVal = hasRetVal_Last
'
'                     hasRetVal_Last = hasRetVal
'
'                   ' Clear value on last cond
'                     If isLastCond Then hasRetVal_Last = Empty
'
'
'      '               FasStack.Current = make_ITEM(TokenOpen("cond") & Condition)
'
'                  End If ' isCond
'=================================================================
   
   '               If Not isLastCond Then
   '
   '                  InterpretStream_rek FileStream, _
   '                        Target, .Position + param(1) + 1, _
   '                        level + 1, 0
   '
   '
   '                  asd
   '                  FasCmdline.Interpreted = FasStack.pop & _
   '                              IIf(level = 0, _
   '                                  " )", _
   '                                  "")
   '
   '               Else
   '
   '                  .Move 1 + 1 + 4 'Skip 01_Push <nil>; 57_Goto
   '
   '               End If
           
        


         Case &H6A: 'AND
            
            Param.add make_INT(.int32Sig, 4)
            branchTarget = .Position + Param(1)
            
            DisASM_short = "AND"
            DisASM = "and_If (" & FasStack.Current & ") .pop Else goto " & OffToStr(branchTarget)
            
            
            InterpretStream_output DisASM, DisASM_short, Description

            'Label_Push CStr(tmp), "and"
'            If ANDDest = branchTarget Then
'              'Next/inner andif
'               InterpretStream_rek FileStream, Target, .Position + param(1), level + 1, -1, branchTarget, InterpretStream_AND_6A_Helper
'               InterpretStream_AND_6A_Helper = InterpretStream_AND_6A_Helper + 1
'            Else
'              'First/ outer andif
'               InterpretStream_AND_6A_Helper = 1
'               InterpretStream_rek FileStream, Target, .Position + param(1), level + 1, -1, branchTarget, InterpretStream_AND_6A_Helper
'             '  FasStack.ESP = FasStack.ESP + InterpretStream_AND_6A_Helper - 1
'             '  FasStack.push "(and " & Join(FasStack.popArray(InterpretStream_AND_6A_Helper)) & ")"
'               InterpretStream_AND_6A_Helper = 1
''               disasm = GetIndent(level) &  "-------"
'            End If
            GoTo NoOutput
            
         Case &HF 'fsl - together with 0d_if
            Param.add make_INT(.int16, 2)
            DisASM = "(FSL 0xd_if related)  jump to " & OffToStr(Param(1) + .Position)
            InterpretStream_output DisASM, DisASM_short, Description
            
            If Param(1) > 0 Then
               InterpretStream_rek FileStream, Target, .Position + Param(1), level + 1
            End If
        
            'GoTo noOutput
            
            
         Case &H57 'Goto 32
            DisASM_short = "GOTO"
            
            Dim GotoTarget
            Set GotoTarget = make_INT(.int32Sig, 4)
            
            Param.add GotoTarget
            
            GotoTarget = GotoTarget + .Position
            DisASM = "goto " & OffToStr(GotoTarget)
            
            
         ' Defun Error Handler
            Dim IsErrorHandler As Boolean
            IsErrorHandler = IsNextCommand(FileStream, &H14) 'Defun
            
            If IsErrorHandler Then
            
             ' Output Line
               DisASM = DisASM & " Nested defun..."
               InterpretStream_output DisASM, DisASM_short, Description

               
               Dim ErrorHandler
               ErrorHandler = Brancher.DoBlock(GotoTarget, level + 2)
              
             ' skip error Function load
'    $43   09  0025  pu_Item   push <Func> *ERROR* Modul:_FAS0, Offs: $19 [25]
'    $46   06  0020  setq      *ERROR* =
'    $49
               Dim DefunStart
               DefunStart = .Position + (1 + 2) + (1 + 2)
               Debug.Assert IsNextCommand(FileStream, &H9)  'Push Item
               Brancher.DoBlock DefunStart, level + 1, 1
               'Set FunctionToDecompile = Nothing
               'FasStack.popIntoVoid
               
               
             ' set output
               FasCmdline.Interpreted = Join(ErrorHandler, vbCrLf)
               
               'GoTo NoOutput
            End If

            
            


            
         Case &H21
            Param.add make_INT(.int16, 2)
            DisASM = "end defun ;cleanup " & Param(1) & " vars (" & Param(1) & " " & Join(FasStack.popArray(Param(1) + 2), ",") & ")"
            'select case param1 '[61A49788]
            'case 1..10    5,6 invalid cmd  2 exit
'         Case &H22
'            Stop

         Case &H3E 'Pop and Exit Func if not Zero
            DisASM = "end Defun if " & FasStack.pop & " != ZERO"
            
         Case &H3F 'Pop and Exit Func if  Zero
            DisASM = "end Defun if " & FasStack.pop & " == ZERO"
            
       ' defun) 'FAS
         Case &H16: 'Defun
            DisASM = "end Defun"
            
            bSuppressDecompOutput = m_MVar = 0

            FasCmdline.Interpreted = FasStack.pop & _
                                    IIf(bSuppressDecompOutput, "", _
                                        vbCrLf & TokenClose("defun"))
       '     asd
            
       '
            StopAtOffset = .Position
            
            'in case of assert - Checking if stack makes sence
'            Debug.Assert FasStack.ESP = 0
            
 '           FasStack.ESP = 0 'Clearn stack
         
         
         Case &H1C 'Copy stack to Functions Start/End
            FasCmdline.Interpreted = FasStack.pop
            
            DisASM_short = "I_DONE"
            DisASM = "init done." & vbCrLf & vbCrLf & String(30, "=")
            
            Debug.Assert FasStack.esp = 0
            FasStack.esp = 0
            'ebp+c=0
            '[61A49790]=[[61A49790]+C]
            'pop, reset varbasepionter and something else...
 '           Set m_MVar = m_MVars(1)
 
'         Case &H34
'            DisASM_short = "AND"
'            param.add make_INT(.int8, 1) ' Func Arguments
'            Dim Flags
'                Flags = .int8   'used bits:
'                                  '           bit_0, cleared -> IsACAD Function (for ex C:INIT )
'                                  '           bit_2, set     -> Invalid Function
'                                  '           bit_1
'            param.add make_INT(Flags, 1) ' Flags
'
'            FasStack.push FasStack.pop & " " & Join(FasStack.popArray(param(1)))
'
'          ' Show last 3 bit's
'            Debug.Assert 0 = (Flags And &HFC&) 'Ensure others are cleared
'            DisASM = FasStack.Current & ", " & _
'                                         boolToText(Flags And 1) & ", " & _
'                                         boolToText(Flags \ 4 And 1) & ", " & _
'                                         boolToText(Flags \ 2 And 1) & ", "
         
         'USUBR
         Case &H34, &H35, &H51:
                  
                  Dim IsH34_EVAL As Boolean
                  IsH34_EVAL = FasCmdline.Commando = &H34
                  
                  Dim IsH35 As Boolean
                  IsH35 = FasCmdline.Commando = &H35
                  
                  Dim IsH51_Func As Boolean
                  IsH51_Func = FasCmdline.Commando = &H51
                  
                                                      
                  DisASM_short = Switch( _
                     IsH34_EVAL, "EVAL", _
                     IsH35, "ld_USUBR", _
                     IsH51_Func, "FUNC", _
 _
                     True, "_ERROR" _
                     )
                     
                  Description = "Load User subroutine from stream"
                  
                ' Param #1 - ParamAbove
                  Dim ParamAbove
                  Set ParamAbove = make_INT(.int8, 1) 'Params are above (stored in [ebp+c])
                  Param.add ParamAbove
                  
                ' Param #2 - GVarIdx
                  If IsH35 Or IsH51_Func Then
                     Param.add make_INT(.int16, 2) 'GVarIdx
                  End If
                  
                ' Param #3 - Flags
                  Dim Flags
                  Flags = .int8   'used bits:
                                  '           bit_0, cleared -> IsACAD Function (for ex C:INIT )
                                  '           bit_2, set     -> Invalid Function
                                  '           bit_1
                                  
                ' Param #4 - ??? 51Only
                  Dim IsH51Only
                  If IsH51_Func Then
                     IsH51Only = .int8
                     
                   ' Should be 00
                     Debug.Assert IsH51Only = 0
                  End If
                                  
                                  
                  If IsH34_EVAL Then
                   ' so far it seem that there are all the times all flags set
                     Debug.Assert Flags = 7
                     Param.add boolToText(True)
                     Param.add boolToText(True)
                     Param.add boolToText(True)
                     
                     'delme
                     Debug.Assert Flags = 0
                     
                   ' There is no defun name...
                     Dim DefunName
                     Set DefunName = make_ITEM("")
                     
                   ' ... and always one argument on the top
                     Expression = FasStack.popArray(ParamAbove + 1)
 
                     
                  ElseIf IsH35 Or IsH51_Func Then
                     Debug.Assert 0 = (Flags And &HFC&)
                     Dim b_0_IsWhatever As Boolean
                     b_0_IsWhatever = Flags And 1
                     Param.add boolToText(b_0_IsWhatever)
   '                  Debug.Assert b_0_IsWhatever = True '(provide 'VILLFLB) (providedp 'stampedApp)
   
   '                  Debug.Assert b_0_IsWhatever = False
   
                     
                     Dim b_2_IsIvalidFunc As Boolean
                     b_2_IsIvalidFunc = Flags And 4
                     Param.add boolToText(b_2_IsIvalidFunc)
                     Debug.Assert b_2_IsIvalidFunc = False
                     
                     
                     Dim b_1_IsNotACAD_BuildIn As Boolean
                     b_1_IsNotACAD_BuildIn = Flags And 2
                     Param.add boolToText(b_1_IsNotACAD_BuildIn)
                     'Debug.Assert b_1_IsNotACAD_BuildIn = True
                     
   
                     
                     'Debug.Assert (Flags = 3) Or (Flags = 1)
                     
                     'Debug.Assert .Position <> 2735
                     
   '                  On Error Resume Next
                  
                     Set DefunName = MVars(Param(2))
                     Expression = FasStack.popArray(ParamAbove)
                     
                  End If
                     
                  
               '-- Fetch Done --
               
               
               
               Dim PushThis
               
               DisASM = IIf(Not (b_1_IsNotACAD_BuildIn), "ACAD-Func ", "") & _
                  DefunName & " " & _
                  ParamAbove & IIf(IsH34_EVAL, "(+1)", "") & _
                  " Params are above..."
                  'FasStack.ESP = FasStack.ESP - ParamAbove
                  'FasStack.push "Retval of " & DefunName


               Select Case DefunName
               
               Case "vl-ACAD-defun":                  Debug.Assert IsH35

                  'PushThis = TokenFull(DefunName, Expression)
                  Set PushThis = make_ITEM("", True)
                  
               Case "ads-cmd":                        Debug.Assert IsH51_Func

               
               '  Sample
               '   $86   9  28         1 pu_Item   push "_.UNDO"
               '   $89  51  1 27 + - - 1 FUNC      ACAD-Func ads-cmd 1 Params are above...
               '   $8F   A             0 Pop       pop dummy (decrease stack)                               (Command "_.UNDO")
               '
               '   $90   9  26         1 pu_Item   push "_E"
               '   $93  51  1 27 + - - 1 FUNC      ACAD-Func ads-cmd 1 Params are above...
               '   $99   A             0 Pop       pop dummy (decrease stack)                               (Command "_E")
               '
               '   $9A  35  0 25 + - + 1 ld_USUBR  AI_UNDO_OFF 0 Params are above...
               '   $9F   A             0 Pop       pop dummy (decrease stack)
                  
                  Dim lastGoodPos
                  lastGoodPos = .Position
                     Dim commanditems
                     For commanditems = 0 To &HFFFFFFF
   
                        If Not And_( _
                           .int8 = &HA, _
                           .int8 _
                        ) Then Exit For
                        
                        Dim CommandIdx
                        CommandIdx = .int16
                        
                        If Not And_( _
                           .int8 = &H51, _
                           .int8 = &H1, _
                           .int16 = Param(2), _
                           .int8 = Flags, _
                           .int8 = 0 _
                        ) Then Exit For
                        lastGoodPos = .Position

                        
                        Dim CommandToAdd
                        Set CommandToAdd = MVars(CommandIdx)
                        ArrayAdd Expression, CommandToAdd
                        
                     Next
                  .Position = lastGoodPos
                  
                  Set PushThis = make_ITEM(TokenFull("Command", Expression))

               Case "collect-init", "collect-next", "collect-release"
                  Debug.Assert IsH51_Func
                  
                  'PushThis = TokenFull(DefunName, Expression)
                  Set PushThis = make_ITEM("", True)
                  
                  If DefunName = "collect-init" Then
                     Set Brancher.VlaxFor_item = make_ITEM(Expression(1))
                  End If
                  
                  
              
               Case "string-resource"
                  Debug.Assert IsH51_Func
                  
                  
                  Dim DllName, ResID, ResText
                  DllName = Expression(1)
                  ResID = Expression(2)
                  ResText = Quote( _
                     LoadResStringDll(ResID, _
                                             App.Path & "\" & DeQuote(DllName) _
                           ))
                  DisASM = ResText & " '" & DefunName & "' from " & DllName
                  Set PushThis = make_ITEM(ResText) '& "<-" & tmpstr2
                  
               
               Case "_al-bind-alist"
               
                  Debug.Assert IsH51_Func
                  
                ' Patch Decompiled lines (LispFileData)
                  Dim Defun_Def
                  Defun_Def = FrmMain.LispFileData.pop
                  
                  If InStr(Defun_Def, ")") Then
                    'Defun xyz (x1, x2 / a1, a2)
                     Defun_Def = Replace(Defun_Def, ")", Replace(JoinToText(Expression), "'(", " / "))
                  Else
                    'Defun xyz (/ a1, a2)
                     Defun_Def = Defun_Def + Replace(JoinToText(Expression), "'(", "(/ ")
                  End If

                  
                  'FasStack.push "" 'tmpstr1
                  Set PushThis = make_ITEM(Defun_Def)
                  'FrmMain.LispFileData.push tmpstr1
                     
               Case Else
               
                  
                  
                ' defun-Handler#2a: for possible Lambda
                  If Not FunctionToDecompile Is Nothing Then


                   ' Exclude C:XXX Functions as Lambda
                     FunctionToDecompile.isLambda = b_1_IsNotACAD_BuildIn
                     
                 End If
                 
                 If (Not FunctionToDecompile Is Nothing) And b_1_IsNotACAD_BuildIn Then
                 
                   ' Just checking - since the name of lambda functions is always "-lambda-"
                     Debug.Assert FunctionToDecompile.FuncName = "-lambda-"
                     
                     
                   ' ACAD_BuildIn FunctionName starts with "C:"
                     Dim bIsNotACAD_BuildInByName
                     bIsNotACAD_BuildInByName = FunctionToDecompile.FuncName Like "C:*"
                     
                   ' Test if C:XXX Functions and the Flag always goes together
                     Debug.Assert (Not b_1_IsNotACAD_BuildIn = bIsNotACAD_BuildInByName)
                     
                     
                   ' Keep Parathese open for the following lambda function...
                     PushThis = TokenOpen(DefunName) & JoinToText(Expression)
                     Inc LambdaOpenParatheseCount
                  Else
                     PushThis = TokenFull(DefunName, Expression)
                  End If
                  asd
                  If b_0_IsWhatever = False Then 'Not CBool(m_MVar)
                     Debug.Print "b_0_IsWhatever: ", PushThis
                  End If


                ' Brancher.ForEach_item And _

                'Not really good for fsl
                  'If Not b_1_IsNotACAD_BuildIn Then
                   ' supress  (vl-ACAD-defun C:XXX) output of ACAD_BuildIn Functions
                 '    Debug.Print "Wasting: " & DefunName
'                '     Stop
                 '    PushThis = JoinToText(Expression)
                 ' End If
                  
                  Set PushThis = make_ITEM(PushThis)
                  
              End Select

               '--------------------
                                 
                  
               FasStack.push PushThis
               
               
                  
                  
''Convert to funcptr
'         Case &H51
'              DisASM_short = "FUNC"
'              param.add make_INT(.int8, 1) 'gets overwrite with 45 or 5A
'              param.add make_INT(.int16, 2)  'GVarIdx    ->long(lo)
'              param.add make_INT(.int16, 2)  'test if 00 ->long(hi)
'              '"<- Convert to funcptr and ebx--"
'
'              Dim FuncName
'              FuncName = MVars(param(2))
'              Select Case FuncName
'
'               Case "string-resource"
''                  Dim tmpstr1$, tmpstr2$, tmpstr3$
'                  tmpstr1 = FasStack.pop
'                  tmpstr2 = FasStack.pop
'                  tmpstr3 = """" & _
'                     LoadResStringDll(tmpstr1, _
'                                             App.Path & "\" & Split(tmpstr2, """")(1) _
'                           ) & """"
'                  DisASM = tmpstr3 & " '" & FuncName & "' from " & tmpstr2
'                  FasStack.push tmpstr3 '& "<-" & tmpstr2
'
'               Case Else
'                  DisASM = FuncName & " " & param(1) & " Params are above <- Convert to funcptr"
'                  'FasStack.ESP = FasStack.ESP - param(1)
'                  'FasStack.push "Retval of " & MVars(param(2))
'                  tmpstr2 = Join(FasStack.popArray(param(1)))
'
'
'                ' Patch Decompiled lines (LispFileData)
'                  tmpstr1 = FrmMain.LispFileData.pop
'
'                  If InStr(tmpstr1, ")") Then
'                    'Defun xyz (x1, x2 / a1, a2)
'                     tmpstr1 = Replace(tmpstr1, ")", Replace(tmpstr2, "'(", " / "))
'                  Else
'                    'Defun xyz (/ a1, a2)
'                     tmpstr1 = tmpstr1 + Replace(tmpstr2, "'(", "(/ ")
'                  End If
'
'                  If FuncName = "_al-bind-alist" Then
'                     'FasStack.push "" 'tmpstr1
'                     FasStack.push tmpstr1
'                     'FrmMain.LispFileData.push tmpstr1
'
'                  Else
'                     FasStack.push "(" & FuncName & " " & _
'                          tmpstr2 & ")"
'
'                  End If
'
'
'              End Select

            
         Case &H45 'Function in Ram (Converted)
            Param.add make_INT(.int8, 1)
            Param.add make_INT(.int32, 4)
            DisASM = "Memdumpfunc !!! Call " & Param(2) & " (" & Join(FasStack.popArray(Param(1))) & ")"
            FasStack.push DisASM
         
         Case &H5A  'Function in Ram (Converted)
            Param.add make_INT(.int8, 1) 'Number of Stack elements; 0..3 valid else "error in XNCALL"
            Param.add make_INT(.int32, 4) '->fctptr "..\KERN\FNS\XRUN.CPP"

            DisASM = "Memdumpfunc !!! Call VL.ARX!" & Param(2) & " (" & Join(FasStack.popArray(Param(1))) & ")"
            FasStack.push DisASM
         
         
         
         Case &H5F
            DisASM_short = "CaByOffs"
            
          ' Param #1 - ParamAbove
            Set ParamAbove = make_INT(.int8, 1) 'Params are above (stored in [ebp+c])
            Param.add ParamAbove
            
          ' Param #2 - Function Offset
            Dim FuncOffset
            FuncOffset = make_INT(.int32, 4)
            Param.add FuncOffset
            
          ' Fetch Params
            Dim Params
            Params = FasStack.popArray(Param(1))
            
          ' Fetch Function
            DisASM = ""
            On Error Resume Next
            DisASM = FasFunctionnames(CStr(FuncOffset))
            On Error GoTo InterpretStream_rek_err
 
          ' Output
            FasStack.push make_USUBR(TokenFull(DisASM, Params))
            DisASM = "CallByOffset @" & OffToStr(FuncOffset) & " : " & DisASM & " (" & Join(Params) & ")"
            
         
         Case &H53
         '<vl.public: virtual short closure::ltid(void)>, vl.closure::ltid
               DisASM = "Xpct-Handler53! " & Join(FasStack.popArray(2))
               FasStack.push "Retval of " & DisASM
             ' disasm = "Reset? " & FasStack.pop & " & " & FasStack.pop
              '[dgcpro::head]
              '<vl.public: virtual short cons::ltid(void)>), vl.cons::ltid

              '"bad argument list tail:"
         
         Case &H54
              'disasm = "unknow delta F"
              DisASM = "Xpct-Handler54! " & Join(FasStack.popArray(2))
              FasStack.push "Retval of " & DisASM
         
         Case &H2E
            Param.add make_INT(.int8, 1) 'flag ?
            DisASM = "StackCallAndJmp " & FasStack.pop & " (" & Join(FasStack.popArray(Param(1)), ",") & ")"
         
         Case &H2F
            Param.add make_INT(.int8, 1)
            Param.add make_INT(.int16, 2)
            DisASM = "CallAndJmp " & MVars(Param(2)) & " (" & Join(FasStack.popArray(Param(1)), ",") & ")"
            
         Case &H60
            
            Param.add make_INT(.int8, 1) '0
            Param.add make_INT(.int32, 4) '1
            
            DisASM = ""
            On Error Resume Next
            DisASM = FasFunctionnames(CStr(Param(2)))
            On Error GoTo InterpretStream_rek_err
 
            DisASM = "jmp2_nopop " & DisASM & " at " & Param(2) & "(" & Join(FasStack.popArray(Param(1)), " ") & ")"
            FasStack.push "Retval of " & DisASM

         Case &H61
            Param.add make_INT(.int8, 1)
            Param.add make_INT(.int32, 4)
            DisASM = "[" & Param(2) & "]"
            On Error Resume Next
            DisASM = FasFunctionnames(CStr(Param(2))) ' param2 is 32 - is that correct -normally it is a int16
            
            Stop 'for Exploring
            
            On Error GoTo InterpretStream_rek_err
 
            DisASM = "continue at " & DisASM & " " & Join(FasStack.popArray(Param(1)), " ") & ")"
            FasStack.push "(" & DisASM & ")"
            
         Case &H18: 'Copy Stack to local Var
            DisASM_short = "iArgs"
            Description = "Init args and local vars with stack"
            
          ' Param #1 - ParamsCount
            Dim ParamsCount
            ParamsCount = .int16
            Param.add make_INT(ParamsCount, 2)
            
            Dim isStartof_ForEach As Boolean
            isStartof_ForEach = Last_FasCmdline.Commando = &H1 'NIL
            
            Dim LVars_
            tmp = (ParamsCount \ 2) * 2 'Align to 2 - to ensure that there are always pairs
            LVars_ = FasStack.popArray(tmp)
            DisASM = "Init ("
            
            If isStartof_ForEach Then
               DisASM = DisASM & "...Start of For Each..."
               Debug.Assert UBound(LVars_) = 2 'sometimes 4
               
               Set Brancher.ForEach_item = LVars_(1)
               asValue Brancher.ForEach_item
               
            Else

               
             ' Maybe triggered if called more than once
             ' happens on 'for each'
               Debug.Assert FasFuncCurrent.Args.count = 0
               
               
               If IsArray(LVars_) Then
               
                  For i = LBound(LVars_) To UBound(LVars_) Step 1
                  
                     If i Mod 2 Then
                       'VarName
                       Dim VarName
                       Set VarName = LVars_(i): asValue LVars_(i)
                       
                     ' Add Arg
                       FasFuncCurrent.Args.add VarName, CStr(i)
                       Debug.Assert FasFuncCurrent.Args(CStr(i)) = VarName
   
                        DisASM = DisASM & VarName
                        FasCmdline.Interpreted = FasCmdline.Interpreted & "  " & VarName
                     Else
                       'Pos LocalVar
                        DisASM = DisASM & "=[L" & LVars_(i) & "]  "
                     End If
                  Next
                  
               End If
               
               
             ' Patch in args
               Dim tmpstr1$, tmpstr2$ ', tmpstr3$
               tmpstr1 = FrmMain.LispFileData.pop
               tmpstr2 = "(" & LTrim$(FasCmdline.Interpreted) & ")"
               
               If InStr(tmpstr1, "defun") Or _
                  InStr(tmpstr1, "lambda") Then
                  
                  tmpstr1 = Replace(tmpstr1, " ( )", "")
                  FasCmdline.Interpreted = tmpstr1 + " " + tmpstr2
               Else
                  FasCmdline.Interpreted = tmpstr2
               End If
               
   
               tmp = Empty
               
            End If
            
            DisASM = DisASM & ") " & (ParamsCount And 1)
           'Debug.Assert (param(1) And 1) = 0
            
         Case &H19:
            Param.add make_INT(.int16, 2)
'            Debug.Assert (param(1) And 1) = 1  'stop when this is 0 - for further exmiation?
            'tmp = param(1) \ 2 + 1
            DisASM = "clear " & Param(1) & " args+vars" '& _
               IIf((param(1) And 1), "", " last func in stream !")
'            FAsStack.ESP = FAsStack.ESP - tmp
                           
'         Case &H1D
'            Stop
         Case &H1E
            Param.add make_INT(.int8, 1)
            DisASM = "??? alpha " & FasStack.pop
         Case &H1F
            Param.add make_INT(.int8, 1)
            DisASM = "??? beta"
'Symbol-Handling Functions
         Case &H23 ' Null and Not is same - Typically
                   'null' is used for lists, and
                   'not' is used for other data types along with some types of control functions.
               DisASM = "Is " & FasStack.Current & " == 0 ?"
               FasStack.push "(NULL " & FasStack.pop & ")"

         Case &H24
            DisASM = "Is " & FasStack.Current & " no list?"
            FasStack.push "(atom " & FasStack.pop & ")"
         Case &H25
            Param.add make_INT(.int8, 1)
            tmp = "LVar[" & Param(1) & "]"
            DisASM = "IsSym (FSL) " & tmp & " & " & FasStack.Current 'ltid
'            FasStack.push tmp
            
'            disasm = "push (cmp " & param(1) & ", " & FasStack.pop
         Case &H26
            tmp = FasStack.pop & " - " & FasStack.pop
            FasStack.push tmp
            DisASM = "push (" & tmp & ")"
            
' === Listmanipulations === (Cons)

'CONStructs memory objects
'A CONS cell is composed of two pointers; (<-non-atomic s-expressions "NATSes" or  pairs)
'the CAR operation extracts the first pointer, and
'the CDR operation extracts the second.

'CAR => Content of Address   Register or first element
'CDR => Content of Decrement Register or second element

       '(car (cons x y)) => x
         Case &H28 'Get first element
            DisASM = "push data element(1.st) of ( " & FasStack.Current & ") ;CAR"
            DisASM_short = "CAR"
            
            FasStack.push make_ITEM(TokenFull("car", FasStack.pop), True)
            'tmp = "list element from (" & FasStack.pop & ")" 'losid::mark(void)'[ebx+08h]
            'disasm = "push " & tmp
            'FasStack.push tmp
            
        '(cdr (cons x y)) => y
         Case &H29  'Get first elements
            DisASM = "push list element(2.nd) of (cons " & FasStack.Current & ") ;CDR"
            DisASM_short = "CDR"

            FasStack.push make_ITEM(TokenFull("cdr", FasStack.pop), True)
            
'            tmp = "next list element from (" & FasStack.pop & ")" 'losid::mark(void) '[ebx+0Ch]
'            disasm = "push " & tmp
'            FasStack.push tmp

         Case &H2A  'Insert element at beginning 'FSL
            DisASM_short = "cons"
            Description = "Pops two elements from stack and pushes a list "
'            If FasStack.ESP > 1 Then
               Set tmp = make_LIST(FasStack.popArray(2))
               tmp.isCons = True
            FasStack.push tmp

'            Else
'               tmp = "Stackerror!"
'            End If
'            'disasm = "defFunc " & Join(tmp, " .. ")
            If IsNumeric(tmp.data(1)) And IsNumeric(tmp.data(2)) Then
               DisASM = "(Insert " & OffToStr(CLng(tmp.data(1))) & _
                           "  at beginning of " & OffToStr(CLng(tmp.data(2)))
            Else
               DisASM = "(Insert " & tmp.data(1) & "  at beginning of " & tmp.data(2)
            End If
'            Dim FasFunction As FasFunction
'            Set FasFunction = New FasFunction
'
'            FasFunction.Startoffset = tmp(LBound(tmp))
'            FasFunction.Endoffset = tmp(UBound(tmp))
'
'            On Error Resume Next
'            FasFunctionnames.Remove CStr(FasFunction.Startoffset)
'            On Error GoTo InterpretStream_rek_err
'            FasFunctionnames.Add FasFunction, CStr(FasFunction.Startoffset)
'
            'Typ [cons]
            '[esp-2]=funct(esp-2,esp-1)
            
         Case &H2C
            Param.add make_INT(.int8, 1) 'index
            tmp = FasStack.pop & " [" & Param(1) & "]" '& MVars(param(1))  '"losid_subtype_p"
            FasStack.push (tmp)
            DisASM = tmp
            
         Case &H2D
            Param.add make_INT(.int8, 1) 'index
            tmp = FasStack.pop & " = " & FasStack.pop & " [" & Param(1) & "]" '& MVars(param(1)) 'losid_subtype_p
            FasStack.push tmp
            DisASM = tmp
            
            
        'Load_INT8
         Case &H32:
            DisASM_short = "Ld_INT8"
            Description = "push signed 8 Bit Integer from stream => stack"
            
            Param.add make_INT(.int8Sig, 1)
            
            DisASM = "push  " & Format(Param(1), "00")
            FasStack.push Param(1)

        'Load_INT32
         Case &H33:
            DisASM_short = "Ld_INT32"
            Description = "push signed 32 Bit Integer from stream => stack"

            Param.add make_INT(.int32, 4) 'var1=v+v
            
          '' Speciality why this value was not encoded using the INT8 opcode?
          '  Debug.Assert param(1) > &HFF&

            'push .int32
            DisASM = "push " & Format(Param(1), "00000000")
            FasStack.push Param(1)
         
         
        'Load_STRs
         Case &H55
            DisASM_short = "Ld_STR"
            Description = "push Strings[lnode] from stream => stack"

         
'           param.add make_INT(.int16, 2)sig 'Number of strings that follow
           
           'Read all Strings
            For i = 1 To .int16 'param(1)
'               param.add make_INT(.int16, 2) 'Stringlength
                                                                            '=Last Parameter
               Set tmp = make_STR(.FixedString(.int16))
'               tmp = .fixedString(param(param.Count))
               Param.add tmp
             
             ' Add String to Stack ( will be loaded into StringTable later)
               FasStack.push tmp

            Next
            
         
        'Load_SYMs
         Case &H56, &H5B 'Create Stringtable
            'FSL 56
            DisASM_short = "Ld_SYM"        'SUBR ?5B
            Description = "push Symbols from stream => stack"
            
            DisASM = "Push&load Symbols"
            
            'Dim isFas As Boolean
            isFas = FasCmdline.Commando = &H5B
 
            Do
             ' get symbol string
               Set tmp = make_SYM(.zeroString)
             
             ' is the string the "Terminatorstring" (=Zerolength)
               If Len(tmp) = 0 Then Exit Do
             
             ' store symbolString
               Param.add tmp
               
               
'             ' FAS command are stored in upcase
'             ' convert them to lower case with is more common
'               If isFas Then
'                  ToLowerCase tmp
'               End If
               
             ' Add String to StringTable
               FasStack.push tmp
            Loop While True
         
         
         Case &H40
            Param.add make_INT(.int16, 2) '+0
            Param.add make_INT(.int32, 4) '+2
            DisASM = "??? memMov sth."
            
            FasStack.push DisASM
         
         Case &H37
            Param.add make_INT(.int16, 2) 'int16sig
            tmp = "ListObject B with " & Param(1) & " elements"
            DisASM = "push " & tmp '[cons]
            'pop * number of param(1); push FuncListObject
'            FAsStack.ESP = FAsStack.ESP - param(1)
            FasStack.push "<list( " & Join(FasStack.popArray(Param(1)), ",") & ")>"
            
            
         Case &H38:
               DisASM = "Convert last element on stack"
               
         'Ld_REAL
         Case &H3B
              
              Param.add make_REAL(.zeroString)
              FasStack.push Param(1)
              DisASM_short = "Ld_REAL"
              DisASM = "push 'Floating-point' number string for stream => stack"
         
       ' Ld_LIST
         Case &H39   'Choose Array for store strings
            Param.add make_INT(.int16, 2)
            FasStack.push make_LIST(FasStack.popArray(Param(1)))
            
            DisASM_short = "Ld_LIST"
            DisASM = "Combines " & Param(1) & " elements on the stack to a LIST"
            
       
        'Define Functions
         Case &H3A
            DisASM_short = "ld_USUBR"
            Description = " Load user subroutiune from stream => GVar; 3x Stack"
            
            Set tmp = FasStack.pop     ' Name
            p1_LVrs = FasStack.pop     ' StartOff
            p2_Args = FasStack.pop     ' Module
           
           ' Dim FasFunction As FasFunction
            Set FasFunction = New FasFunction
            
            FasFunction.Startoffset = p1_LVrs
            
            FasFunction.ModulId = p2_Args
            Set FasFunction.ModulStream = FasFunctionData
            
'          ' set to end of Eond of stream till we now it better.
'            FasFunction.Endoffset = FasFunctionData.Length
'            Dim Endoffset&
'            Endoffset
'
'          ' check bounds and set endoffset of the last added function
'          '
'          ' assumption: functions are in Order and are added in ascending order
'          ' (^- at the moment this works, but may neew to be improved)
'            Dim LastFunction As FasFunction
'            Set LastFunction = FasFunctionnames(FasFunctionnames.count)
'
'            Debug.Assert LastFunction.Startoffset = FasFunction.Startoffset
'
          
            
' DIRTY-Code: the moduleparameter (p2_Args) is not used -
'             We just put all the functionoffset into collection
'            there can be offset-collisions'
'so sometimes we have to ".remove" before ".add"

            On Error Resume Next
            
           'Add new function with key offset
            Dim FasFunctionnameItem
            FasFunctionnames.add FasFunction, CStr(p1_LVrs)
            If Err = 457 Then
               If FasFunction.FuncName <> Empty Then
               '... but only if it contains a Name
                  Stop
                  FasFunctionnames.Remove p1_LVrs
                  'And again
                   FasFunctionnames.add FasFunction, p1_LVrs
               Else
               
               End If
            ElseIf Err <> 0 Then
               'other err
               Stop
            End If
            
            
            
            DisASM = tmp & _
                        " Modul:" & p2_Args & ", " & _
                        "Offs: " & OffToStr(p1_LVrs) & " [" & p1_LVrs & "]" '& _
                        vbCrLf  '[xsubr]
                        
            FasFunctionnames.item(CStr(p1_LVrs)) = tmp
            
            Set tmp2 = make_USUBR(Func_TypePrefix & DisASM)
            tmp2.Name = tmp
            tmp2.Start = p1_LVrs
            
            On Error GoTo InterpretStream_rek_err
            
            FasStack.push tmp2  '& tmp
            
            DisASM = "def_Func " & DisASM
            
         Case &H59
         
          ' Tempfixjust to embalance stack
          ' not comfirmed by code analysis
            tmp = FasStack.popArray(9)
            DisASM = "SetupErrorHandler "
            DisASM_short = "hERROR"
            
            Dim Logger2 As New clsStrCat
            Logger2.Clear 'since we my come here several times
            
            Logger2.Concat vbCrLf
            For i = 9 To 1 Step -1
               
               'Indent with 4 Tab's
               Logger2.Concat String(4, vbTab) & ";; "
               Logger2.Concat "#" & BlockAlign_r(i, 1) & " = " & tmp(i)
    '           Logger.Concat IIf(i Mod 2, vbCrLf, vbTab & vbTab)
               Logger2.Concat vbCrLf

            Next
            
            FasCmdline.Interpreted = DisASM & Logger2.value
            
            DisASM = DisASM & Join(tmp)
            
            FasStack.push make_ITEM(FasCmdline.Interpreted)   '";; " & disasm
            
            
         'Stream init(
         Case &H43
            '  Stack:=moduleOrStream
            '  p1:=StartIdx
            '  p2:=NumElementsToPopForStack
            ')
         'It is very well for getting Strings and Functions
         'It is called 3 Times
         '  #1 => this stream: Strings
         '  #2 => functionStream: Strings
         '  #3 => this stream: Append function definition
         ' After initialiation is finished decompilation can start
         
            Static case_43_counter  'helped to decide which stringtable to use
            Inc case_43_counter
            Debug.Assert case_43_counter <= 3 'so far this is called only 3 times during initialisation
         
         
            DisASM_short = "iVars"
            DisASM = "pops elements from stack => Var Storage; 2x param's, 1x stack"

            Dim St_init, VarPos
            
          ' Param #1 - VarPos
            VarPos = .int16
            Param.add make_INT(VarPos, 2)
            
          ' Param #2 - St_init
            St_init = .int16
            Param.add make_INT(St_init, 2)

            
          ' Stack -local_ModuleID
            Dim local_ModuleID
            Set local_ModuleID = FasStack.Current: FasStack.popIntoVoid
            
          ' Known values so far
'            Debug.Assert local_ModuleID = 0 Or _
'                         TypeOf local_ModuleID Is T_NIL

            DisASM = "St_init " & St_init & " Stackitem => " & _
                     " VarPos " & VarPos & " @" & _
                       IIf(TypeOf local_ModuleID Is T_NIL, _
                           "nil", _
                           local_ModuleID _
                        )
                       'IIf(tmp = "nil", "this stream", "FuncStream")
                       
            
          ' Fill ModulVars
            local_ModuleID = IIf(TypeOf local_ModuleID Is T_NIL, 0, 1) 'Choose(case_43_counter, 0, 1, 0)
            
            Dim DoLogging As Boolean
            DoLogging = opt_Log_DumpModulVars
            
          ' Copy from Stack into ModulVar Storage
            Dim Logger As clsStrCat: Set Logger = New clsStrCat
            If DoLogging Then Logger.Concat vbCrLf
            
            
          ' Go from VarPos ... St_init in reverse (since pop it from the stack gives reverse order)
            For i = VarPos + (St_init - 1) To VarPos Step -1
          '
          ' Idea:
          '   implementing a fixed size(size set on init) stack that grows(on push) downwards to 0
          '   (... and offer access via index) we could directly use it
            
            
             ' get item
               On Error Resume Next
               
               Set tmp = FasStack.Current
               If Err Then Err.Clear: Stop: Let tmp = FasStack.Current
               FasStack.popIntoVoid
               
             ' PATCH - Defun Main - don't show Retval (NIL)
             '
               If (local_ModuleID = 0) And (i = St_init - 1) Then
                  Debug.Assert TypeOf tmp Is T_NIL
                  tmp.SupressOutput = True
               End If
               
             ' copy into Modul vars
               MVarsEx(i, local_ModuleID) = tmp

               
             ' Dump Vars into Log
               If DoLogging Then
                 
               ' Indent with 4 Tab's
                 Logger.Concat String(4, vbTab) & ";; "
                 Logger.Concat "#" & BlockAlign_r(i, 8) & " = " & tmp.toText
                 If Err Then
                    Err.Clear
                    Logger.Concat "#" & BlockAlign_r(i, 8) & " = " & tmp
                 End If
      '           Logger.Concat IIf(i Mod 2, vbCrLf, vbTab & vbTab)
                 Logger.Concat vbCrLf
                 
               End If

            Next
            
            If DoLogging Then FasCmdline.Interpreted = Logger.value
            

          ' needs to be after the init loop that pops elements from the stack
            FasStack.push make_ITEM(";; " & DisASM)

           'Create new stringtable only if nessary
'            If (fas43count = 1) Or (fas43count = 3) Then
'               m_MVars.Add m_MVar
'               Set m_MVar = New Collection
'            End If


'=== fixnum operations ===  &H46..&H50
'Binary
         Case &H46 To &H4E
            DisASM_short = Choose(1 + FasCmdline.Commando - &H46, _
               "+", "-", "*", "/", _
               "mod", _
               "<=", ">=", _
               "<", ">" _
               )
               
            'Dim Params
            Params = FasStack.popArray(2)
            
          ' Repeat patch
            Dim bDiscardCommand As Boolean
            bDiscardCommand = Brancher.Repeat_PatchFor_h4B(FasCmdline.Commando, Params(2))
            
            If bDiscardCommand Then
               DisASM = Params(1)
            Else
               DisASM = TokenFull(DisASM_short, Join(Params))
            End If
            
            
            FasStack.push DisASM
            
            If bDiscardCommand Then DisASM = DisASM & " '<= xxx 0' was DISCARDED since a 'Repeat' follows"

            
         
'unary
         Case &H4F, &H50
            DisASM_short = IIf(FasCmdline.Commando = &H4F, _
                           "1+", "1-")

            DisASM = TokenFull(DisASM_short, FasStack.Current)
            FasStack.Current = DisASM
         
         
         
'=== Debug stuff ====
         Case &H65 'e
            DisASM = "TraceIn? " & FasStack.pop
         Case &H66 'f
            '100%
            DisASM = "TraceOut? [61A4974C]=[61A4974C]-4"
         
'nop
         Case &H20, &H62, &H63 ' ',  'b', 'c'
              DisASM = "nop_" & Chr(FasCmdline.Commando)
'invalid
         Case &H11, &H12, &H13, &H1D, &H22, &H27, &H2B, _
              &H30, &H31, &H36, &H41, &H42, &H44, &H52, &H58
            DisASM = "Stop_" & Hex(FasCmdline.Commando) & " Processing this stream"
            
            
         Case Else
            DisASM = "Invalid commando"
            
            Dim Errmsg
            Errmsg = "FasFile::InterpretStream_rek Error: Invalid fas commando" & vbCrLf
            Errmsg = Errmsg & "      @" & OffToStr(FasCmdline.Position) & "  cmd: 0x" & H8(FasCmdline.Commando)
            FrmMain.AddtoLog CStr(Errmsg)
            #If DoDebug = 1 Then
            
                  Dim retval&
                  retval = MsgBox(Errmsg & vbCrLf & _
                                     "Start Debugging? (Cancel to stop at all / No to Continue)", _
                                     vbExclamation + vbYesNoCancel + vbDefaultButton2)
                  If vbYes = retval Then
                     Stop
                     Resume
            
                  '            Stop 'due to invalid commando
                  'DevNotes:
                  'Enter/run  >asd< in VB6 command promt to scroll to latest entry
                  'Enter qw to operated FasDisasm in pause mode - hit ctrl+Pause and drag the yellow IP arrow
                  ' out of the DoEvents infintity loop
                  '-> use Hex Workshop to examine for propably FileDataCorruption
                     
                     
                  ElseIf vbCancel = retval Then
                  
                     Err.Raise ERR_GUI_CANCEL
                     
                  End If
            #End If
            
         End Select
         
         
         InterpretStream_output DisASM, DisASM_short, Description
         
         If FrmMain.Chk_Cancel Then
            Err.Raise ERR_GUI_CANCEL, "", tmp
         End If


       ' defun-Handler#3: Decompile lambda or defun
       ' On pop or end defun
         If FasCmdline.Commando = &HA Or _
            FasCmdline.Commando = &H16 Then
            
            If Not FunctionToDecompile Is Nothing Then
            
             ' Decompile Rekursive
               m_MVar = 1              ' Set Modul to 1
               
               'Debug.Print "Decompiling " & FunctionToDecompile.Name, "@" & FunctionToDecompile.Start

               FasFunctionData.Position = FunctionToDecompile.Startoffset
               InterpretStream_rek FasFunctionData, Target
               
               m_MVar = 0              ' Set Modul to 0
           
               If FunctionToDecompile.isLambda Then
                ' Close outer Lambda statement
                  For i = 1 To LambdaOpenParatheseCount
                     FrmMain.LispFileData.Current = FrmMain.LispFileData.Current & TokenClose
                  Next
               End If
               
               Set FunctionToDecompile = Nothing
            End If
            
         End If
        
         
NoOutput:
         
       ' Save Stack state / at the moment just for debugging purpose
       ' comment out to free some memory
         FasCmdline.Stack_Pointer_After = FasStack.esp
         
         On Error Resume Next
         If IsObject(FasStack.Current) Then
            Set FasCmdline.Stack_After = FasStack.Current
         Else
            Let FasCmdline.Stack_After = FasStack.Current
         End If

         
         If Err Then
            FasCmdline.Stack_After = IIf(Err = ERR_STACK_IS_EMPTY, _
               "<Stack is Empty>", _
               "" _
            )
         End If:         On Error GoTo InterpretStream_rek_err
         

         Set Last_FasCmdline = FasCmdline
         
      Loop
   
'   If FasStack.ESP > StartEsp Then
'      FasCmdline.Interpreted = FasStack.pop
'      InterpretStream_output DisASM, DisASM_short, Description
'   End If

'   when uncomment change 'IF ' which needs stack
'   FasStack.esp = StartEsp


If .EOS Then
   FrmMain.AddtoLog ("Reached End Of Stream at " & OffToStr(.Position) & " File: " & .Filename)
End If

End With

Err.Clear
InterpretStream_rek_err:

'======================================================================================
'                           E R R O R   H A N D L E R


Dim ErrorCounter
If ErrorCounter > 1000 Then On Error GoTo 0

Select Case Err

   Case 0
   
   Case 1001 + vbObjectError 'pop not possible
      Inc ErrorCounter
      
     #If DoDebug = 1 Then
      
            If vbYes = MsgBox("ERROR: pop not possible" & vbCrLf & _
                               "Stop processing ?", _
                               vbExclamation + vbYesNo + vbDefaultButton2) Then Stop
      #End If

      
      Resume Next
'   Case 9 'Index außerhalb des gültigen Bereichs
'      Resume Next
   Case &H800403E9
   
   Case ERR_GUI_CANCEL
      Dim Err_Description$
      Err_Description = Err.Description
      
      On Error GoTo 0
      Err.Raise ERR_GUI_CANCEL, , Err_Description
      
   Case Else
      Inc ErrorCounter
      Errmsg = "FasFile::InterpretStream_rek Error: " & Err.Description & vbCrLf
      Errmsg = Errmsg & "      @" & OffToStr(FasCmdline.Position) & "  cmd: 0x" & H8(FasCmdline.Commando)
      FrmMain.AddtoLog CStr(Errmsg)
      
      Dim defaultButton&
      #If DoDebug = 1 Then
         defaultButton = vbDefaultButton1
      #Else
         defaultButton = vbDefaultButton2
      #End If
      
            'Dim retval&
asd
            retval = MsgBox(Errmsg & vbCrLf & _
                               "Start Debugging? (Cancel to stop at all / No to Continue)", _
                               vbExclamation + vbYesNoCancel + defaultButton)
            If vbYes = retval Then
               Stop
               Resume
            ElseIf vbCancel = retval Then
               Err.Raise ERR_GUI_CANCEL
            End If
      
      Resume Next
      Err.Raise Err, , Err.Description
End Select


End Sub

Private Sub getStreamData(Datalength, INPUT_NumberOfVars, INPUT_NewStream) 'As FileStream)
   Dim StreamTerminatorChar
   Dim tmpChar$
   With File

    
   
   ' Get FunctionStreamVars
     tmpChar = SkipWhiteSpaceEx
     If IsNumeric(tmpChar) = False Then Err.Raise vbObjectError, , "Invalid File Format - Could not get Number of StreamVars."
     
     Do While IsNumeric(tmpChar) And (.EOS = False)
           INPUT_NumberOfVars = INPUT_NumberOfVars & tmpChar
           tmpChar = .char
     Loop
     
     
   ' Get StreamBeginChar (and store it in 'StreamTerminatorChar')
     tmpChar = Skipwhitespace
     If tmpChar = "!" Then
        StreamTerminatorChar = .char
     Else
        StreamTerminatorChar = tmpChar
     End If
  
  
  ' --- Store start of code ---
    CodeStart = .Position
    
    Dim InputFilename As New ClsFilename
    InputFilename.mvarFileName = INPUT_NewStream.Filename
    
    FrmMain.Panel_Detail = "getStreamData " & _
      "0x" & H32(CodeStart) & " => " & InputFilename.NameWithExt

    
  ' --- Fill INPUT_NewStream with data from FILE ---
    On Error Resume Next
    INPUT_NewStream.MemOnlyMode = True '*.res

   
   .Position = CodeStart

    Dim tmpBuffer
    tmpBuffer = .FixedString(Datalength)
    INPUT_NewStream.FixedString(-1) = tmpBuffer
    INPUT_NewStream.Position = 0
    If Err = ERR_OPENFILE Then
         
         Stop

         Set INPUT_NewStream = New StringReader
         INPUT_NewStream.FixedString(-1) = tmpBuffer
         INPUT_NewStream.Position = 0

    End If
'    If Err Then
'    On Error GoTo 0
'       'use file directly (in case we're 'Out of Memory')
'       Set INPUT_NewStream = File
'
'       File.Move (Datalength)
'    End If
    
 

  ' Get next Char - if Char=StreamBeginChar the Stream is uncrypted _
                  - if not it's the length of the following key
    tmpChar = .char
    
  ' --- Decryption ---
    
  ' Set FileEncrypted flag
    Dim IsEncrypted As Boolean
    IsEncrypted = (tmpChar <> StreamTerminatorChar)

    If IsEncrypted Then
       
     ' Set & Check Keylength
       Dim keylength&
       keylength& = Asc(tmpChar)
       
       
       ' 6+15 => 21 so usual Keylen is from 22 .. 27 chars
       If keylength >= &H80 Then Err.Raise vbObjectError, , "crunch password too long. - Keylength is bigger than 128 Byte! Processing canceled."
       '(assert (>= (strlen LVar[0]) 256) "crunch password too long:" LVar[0])

       On Error Resume Next

       KeyStart = .Position
      
      ' Create FasKey & Fill it with data
        Set FasKey = New FileStream
        FasKey.create .Filename & ".key", True, CreateFilesTemporary
        FasKey.FixedString(-1) = .FixedString(keylength)
        If Err = ERR_OPENFILE Then
            
            Stop
            
            Set FasKey = New StringReader
             .Position = KeyStart
            FasKey.FixedString(-1) = .FixedString(keylength)
   
        End If
        On Error GoTo 0
        
        tmpChar = .char
        
       Dim isFas As Boolean
       isFas = (tmpChar = "$")
       
       ' at compile time password is generated in AlcGen.fsl via
       'FAS:
       '(strcat (random-string (+ 6 (random& 5) ) ) "\n;fas4 crunch\n;")
       
       If isFas And (FasFile_Version = 4) Then

       
       '(strcat ... "\n;fas4 crunch\n;")
         Const KeyConstMarker = vbLf & ";"
         Const KeyConstSign = KeyConstMarker & "fas4 crunch" & KeyConstMarker
         Const KeyConstSign2 = vbNullChar & vbNullChar & ";Crunched FAS2-FILE" & vbCrLf & ";"
         
         Dim KeyPart_Const
         FasKey.Position = 0
         KeyPart_Const = Right(FasKey.data, Len(KeyConstSign))
         Debug.Assert KeyPart_Const = KeyConstSign
         
         'random-string (+ 6 (random& 5) ...
         Dim KeyPart_Random
         KeyPart_Random = Left(FasKey.data, keylength - Len(KeyPart_Const))
         Debug.Assert RangeCheck(Len(KeyPart_Random), 6 + 5, 6)
         
       Else
       
       'FSL: LocPDB.fsl
       '(random-string (+ 5 (mod LVar[4] 5)))

       End If

        

    End If
    If tmpChar <> StreamTerminatorChar Then Err.Raise vbObjectError, , "StreamBeginChar and StreamTerminatorChar(after key) are not equal! Processing canceled."
 
   End With
 
 
   ' Speedup by reading files into mem
     File.MemOnlyMode = True '*.fas
 
 
    If IsEncrypted Then
            
            
          Dim KeyOld As Long, KeyNew As Long, value As Long
          
        ' Speedup by reading files into mem
          FasKey.MemOnlyMode = True '*.key

           
          FasKey.Position = 0
          KeyOld = FasKey.int8
         
          If KeyOld <> 0 Then
             
             RaiseEvent DecryptingBegin(INPUT_NewStream.Length)
             
             With INPUT_NewStream
             
              '' Now we can writeback the data without seekback
              ' .DisableMoveOnRead = True
           
              ' Create Filestream to write uncrypted data in FasResourceData
                Dim FasResourceDataWrite As FileStream
                Set FasResourceDataWrite = New FileStream
                FasResourceDataWrite.create (.Filename)
                FasResourceDataWrite.MemOnlyMode = True

              
              
              ' Set filepointer at begin of crypted data
                File.Position = CodeStart
            
                .Position = 0
                Do Until .EOS
              
                   If FasKey.EOS Then FasKey.Position = 0
                   If (Second(time) Mod 2) = 0 Then
'                     DoEvents
                   End If
                   
                   KeyNew = FasKey.int8
            
                 ' Get Data from INPUT_NewStream
                   value = .int8 Xor KeyNew Xor KeyOld
                   
                   RaiseEvent DecryptingProgress(.Position, value)
                   
                   File.int8 = value
                   
                 ' write uncrypted data to FasResourceData
                   FasResourceDataWrite.int8 = value
                   
                 ' .int8 = value
                   
                   KeyOld = KeyNew
                   
'                   DoEvents
                Loop
                
             FasKey.CloseFile
              
             ' copy changes into INPUT_NewStream
              INPUT_NewStream.data = FasResourceDataWrite.data
              FasResourceDataWrite.CloseFile
              
         End With 'INPUT_NewStream


              
          ' Delete Key
          
          
            Dim bDeleteKey As Boolean
            bDeleteKey = True
            If bDeleteKey Then
                ' move beyond key
                  File.Position = (KeyStart) + keylength
                  
                  Dim overlaydata$
                  overlaydata = File.FixedString(-1)
                  
                ' move to start of key
                 Const SIZEOF_KEYLENGTH As Byte = 1
                 File.Position = KeyStart - SIZEOF_KEYLENGTH
                  
                ' Delete it <- Overwrite Key with Data the follow after the key
                 File.FixedString(-1) = overlaydata
                 
                ' Set End of File
                 File.setEOS
             End If
            
            ' Set StreamTerminatorChar so Lisp will recognize this stream as uncrypted
              File.Position = KeyStart - 1
              File.char = StreamTerminatorChar

' ' optionally: Let the key start with \0 so the encrypting Loop will do nothing
'                File.Position = KeyStart
'                File.int8 = 0
             
             
             
         '  Write back changes into file
            File.WriteChanges
            FasResourceDataWrite.WriteChanges
             
              
             RaiseEvent DecryptingDone
              
          '     IsDecryptingDone = True
              
          '     FasResourceData.position = 0
          '     FasResourceData.fixedString(-1) = File.fixedString
              
         End If
   Else
   
'      FrmMain.AddtoLog "Already decrypted."
   End If
   
End Sub


Private Sub fslInit()

End Sub


Private Sub FSLStreamLoad( _
                           OutStream As FileStream, _
                           Ext As String, _
                           StreamLength&, _
                           StreamVars&, _
                           Offset_Start&)
   With File
         '1. Get FunctionStreamLength
          StreamLength = .getTerminatedString("m")
         
         
       
       
       
       
       
         '2. Create FasFunctionDataStream...
           OutStream.create .Filename & Ext, True, CreateFilesTemporary
         '...& Fill it with data
           getStreamData StreamLength, StreamVars, OutStream

         ' Store for use later in Hexeditor jump
           Offset_Start = CodeStart
           
           .getTerminatedString "#"         ' Skip whitespaces

        

   End With
End Sub
Private Sub FASStreamLoad( _
                           OutStream, _
                           Ext As String, _
                           StreamLength&, _
                           StreamVars&, _
                           Offset_Start&)
                           
   
   With File
   
         '1. Get FunctionStreamLength
           Dim tmpChar$
           tmpChar = SkipWhiteSpaceEx
           If IsNumeric(tmpChar) = False Then Err.Raise vbObjectError, , "Invalid File Format - Could not get Length of FasFunctionStream."
           
           Do While IsNumeric(tmpChar) And (.EOS = False)
                 StreamLength = StreamLength & tmpChar
                 tmpChar = .char
           Loop
         
         '2. Create FasFunctionDataStream...
           OutStream.create .Filename & Ext, True, CreateFilesTemporary
         '...& Fill it with data
           getStreamData StreamLength, StreamVars, OutStream


         ' Store for use later in Hexeditor jump
           Offset_Start = CodeStart


   End With
End Sub



Public Function create(Filename As String) As Boolean
    
   CreateFilesTemporary = FrmMain.Chk_cleanup.value = vbChecked
    
With File
  ' Store Filename
    mvarFileName = Filename
    
    .create Filename
    RaiseEvent initBegin
  
     
  ' --- CheckFileType ---
    
  ' Komprimierte VL-Anwendung(VLX) ?
'    If .FixedString(Len(VLX_FILE_SIGNATURE$)) = VLX_FILE_SIGNATURE$ Then
   Dim VLX_Split As New VLXSpliter
   If VLX_Split.VLX_Split(mvarFileName) Then Err.Raise vbObjectError, , "Files of vlx have been extracted."
                                             '^^^^^^^^ ???!
  'TO All VB-Newbees who 'stopped' here rightclick choose
  'switch\break only at notrecoverable errors

    
  ' Autocad LT-Fas-File ?
    If .FixedString(Len(LTFAS_FILE_SIGNATURE)) = LTFAS_FILE_SIGNATURE Then
       Dim LT_Fas_Format_is_not_supported
       LT_Fas_Format_is_not_supported = "LT-Fas Format is not supported yet."
       MsgBox LT_Fas_Format_is_not_supported, vbExclamation
       Err.Raise vbObjectError, , LT_Fas_Format_is_not_supported
    End If
    .Position = 0
    
    Dim tmpChar$
    tmpChar = SkipWhiteSpaceEx
    
   'Is it a FAS or FSL File?
   ' Samples:
   '  FSL: \n#1Y#271m 26...
   '  FAS:  FAS4-FILE ; Do not change it!\n2513\n143 $

    Dim isFSL As Boolean
    isFSL = WHITESPACETABLE(Asc(tmpChar)) = WS_DASH

    
    If isFSL Then
    
      FrmMain.Panel_Detail = "FSL signature found"
      'FSL_FILE INIT
    ' 0.Signature
       If FSL_FILE_SIGNATURE = .getTerminatedString("#") Then
'          fslInit

         FSLStreamLoad _
                     FasFunctionData, _
                     ".fct", _
                     FunctionStreamLength, _
                     FunctionStreamVars, _
                     Offset_CodeStart
         
         '--------------
  
         FSLStreamLoad _
                     FasResourceData, _
                     ".res", _
                     ResourceStreamLength, _
                     ResourceStreamVars, _
                     Offset_DataStart
         
         On Error Resume Next
         ' Sample 687X23.4.2001
            Dim tmp$
            

            Dim Build$
            Build = .getTerminatedString("X")
            tmp = "vill-build-number   : " & Build
            FrmMain.AddtoLog tmp:            FileLog_Add tmp
            
            
            Dim Build_Date$
            Build_Date = .FixedString(-1)
            tmp = "vill-build-localtime: " & Build_Date
            FrmMain.AddtoLog tmp:            FileLog_Add tmp
            
          ' also set Version for fsl - brancher uses this to decide 32 or 16 bit jumps
            'FasFile_Version = 4
  
       Else
          Err.Raise vbObjectError + 1, "Create", "Invalid FSL_FILE_SIGNATURE. No valid Fsl File"
       End If
    
    Else
    
     'FAS_FILE INIT
     
      FrmMain.Panel_Detail = "FAS signature found"
     
 ' 0. Signature
       Dim FileSig As New clsStrCat
     ' Read all following Alphanumeric chars  (until offset 1024 - to speedup checking invalid files)
       Do While (WHITESPACETABLE(Asc(tmpChar)) = WS_ALPHANUMMERIC) And (.Position < 1024) 'note: tmpchar was initialed above this If-Block
         
         FileSig.Concat tmpChar
         tmpChar = File.char
       Loop
     
     ' rewind by one byte (because of 'tmpchar = File.char')
       File.Move -1
   Select Case FileSig
       Case FAS4_FILE_SIGNATURE:            FasFile_Version = 4
       Case FAS3_FILE_SIGNATURE:            FasFile_Version = 3
       Case FAS2_FILE_SIGNATURE:            FasFile_Version = 2
       Case FAS__FILE_SIGNATURE:            FasFile_Version = 1
       Case Else
         Err.Raise vbObjectError + 1, "Create", "Invalid Fas File. FAS__FILE_SIGNATURE not found."
       End Select
       
               
      FASStreamLoad _
                  FasFunctionData, _
                  ".fct", _
                  FunctionStreamLength, _
                  FunctionStreamVars, _
                  Offset_CodeStart
'----------------------
      
      FASStreamLoad _
                  FasResourceData, _
                  ".res", _
                  ResourceStreamLength, _
                  ResourceStreamVars, _
                  Offset_DataStart

       
    End If
   

            
   Dim overlaydata
   
   overlaydata = .FixedString(-1)
   FrmMain.AddtoLog "Remaining Data @0x" & H32(.Position) & ": " & _
            overlaydata
   
   .CloseFile

  End With 'file
  
  'Close Files since they are loaded into memory
  FasFunctionData.CloseFile
  FasResourceData.CloseFile
  '... at this this point the only file that is still open is the *.txt for logging
  

  RaiseEvent InitDone

  If Not (FrmMain.chk_Decryptonly.value = vbChecked) Then

   ' Start Interpreting
     Me.InterpretFile
  Else
     FrmMain.AddtoLog "Decompiling skipped since 'Decrypt only' option is enabled! "
  End If

End Function

'Private Function GetAlphaNumericstring() As String
'   ' Read all following Alphanummeric chars
'       Dim tmpstr
'       Do While WHITESPACETABLE(Asc(tmpchar)) = WS_ALPHANUMMERIC  'note: tmpchar was initialed above this If-Block
'         tmpstr = tmpstr & tmpchar
'         tmpchar = File.Char
'       Loop
'End Function


Private Function Skipwhitespace() As String
   Do
      Skipwhitespace = File.char
   Loop While WHITESPACETABLE(Asc(Skipwhitespace)) = WS_WHITESPACE
End Function

Private Function SkipWhiteSpaceEx() As String
  
 ' Skip leading whitespaces
   SkipWhiteSpaceEx = Skipwhitespace
     
    ' is first char a ';'
      Do While SkipWhiteSpaceEx = ";"
         
       ' seek for CarrierReturn Chr(&HD) or LineFeed Chr(&HA)
         File.getTerminatedString vbCr, vbLf
'         Do
'            SkipWhiteSpaceEx = File.char
'         Loop Until (SkipWhiteSpaceEx = vbCr) Or (SkipWhiteSpaceEx = vbLf)
         
       ' Skip following Whitespaces
         SkipWhiteSpaceEx = Skipwhitespace
   
      Loop
      
      
   
   
End Function

Private Sub Class_Initialize()
  'load Whitespacetable
   WHITESPACETABLE = LoadResData(101, "VL_WHITESPACE_TABLE")
   
End Sub

Private Function LoadResStringDll(ResID, DllName$) As String
   Dim hModul&, hRes&, Buffer$, size&
'   Static skipWarning As Boolean

   hModul = LoadLibrary(DllName)
   If hModul = 0 Then
'      If skipWarning = False Then
'         MsgBox DllName & " not found. Can not load Messages strings." & vbCrLf & "Note: This message will not be show until you restart the program"
'         skipWarning = True
'      End If
      LoadResStringDll = "<Error when getting Res>"
      FrmMain.AddtoLog DllName & " not found. Can not load Messages string " & ResID
   Else
      Buffer = Space(256)
      size = LoadString(hModul, ResID, Buffer, 256)
      LoadResStringDll = Left(Buffer, size)
   End If
End Function
'
'Public Sub Label_Push(Label$, Command$)
'   Dim adder As clsStrCat
'
'   On Error Resume Next
'   Set adder = Label_Push_col.item(Label)
'   If Err Then
'
'      LevelStack.push level
'
'      Set adder = New clsStrCat
'      adder.Concat Command
'      adder.Concat " " & FasStack.pop
'      Label_Push_col.add adder, Label
'   Else
'      adder.Concat " " & FasStack.pop
'   End If
'
'   level = level + 1
'
'    '& "(" & FasStack.Current & ")
'End Sub

'' Push the Label on FasStack if some is found
'' and Pop Level from LevelStack
'Public Sub Label_Pop(Label$)
'   Dim adder As clsStrCat
'
'   On Error Resume Next
'   Set adder = Label_Push_col.item(Label)
'   If Err = 0 Then
'
'      Dim tmp
'      tmp = "(" & adder.value & ")"
'      FasStack.Current = tmp
'
''      FasCmdline.Interpreted = tmp
'      level = LevelStack.pop
'
'   End If
'End Sub

'Converts string to lower case
' if it is completely in UPPER CASE
Private Sub ToLowerCase(ByRef Text_IO)
      Dim isCompletelyUCase  As Boolean
      isCompletelyUCase = UCase(Text_IO) = Text_IO
      
      If isCompletelyUCase Then
         Text_IO = LCase(Text_IO)
      End If
End Sub

'Check is current location is Goto and returns its destination
Private Function isGoto_0x57(ByRef Goto_Addr__OUT As Long, FileStream)
    With FileStream
         Dim Goto_cmd As Byte
         Goto_cmd = .int8
         
         Dim is_Goto_0x57 As Boolean
         is_Goto_0x57 = Goto_cmd = &H57
         Debug.Assert is_Goto_0x57 = True
         
         Goto_Addr__OUT = .int32Sig
      
    End With
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
