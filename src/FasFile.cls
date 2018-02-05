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

Private File As New FileStream

Private CreateFilesTemporary As Boolean

Private FasFunctionData As New FileStream
Public FasFunctionDataInterpeted As New Collection

Private FasResourceData As New FileStream
Public FasResourceDataInterpeted As New Collection

Private FasKey As New FileStream

Private CodeStart&
Public Offset_CodeStart&
Public Offset_DataStart&

Private KeyStart&
Private IsDecryptingDone As Boolean
Private void As Variant
Private mvarFileName As String 'As New ClsFilename
Private FunctionStreamLength&, FunctionStreamVars&
Private ResourceStreamLength&, ResourceStreamVars&

Private Const WS_CONTROLCHAR = 0      'Chr( 0..8,b,e,f)...
Private Const WS_WHITESPACE = 1 '     Tab,LF,NewLine,CR,1A,Space
Private Const WS_BLACKSLASH = 4 '      \
Private Const WS_PIPE = 5 '            |
Private Const WS_ALPHANUMMERIC = &HA ' $0123456789ABCD…
Private Const WS_LIMITER = &HB '       !"'(),;?´{}~
Private Const WS_DASH = &HF '          #
Private WHITESPACETABLE


Dim Last_FasCmdline As FasCommando
Public FasStringtable As Variant
Private FasStringtables(1) As Variant
Public FasStack As New Stack
Private FasFunctionnames As New Collection

Public Event initBegin()
Public Event InitDone()
Public Event DecryptingBegin(BytesToProgress&)
Public Event DecryptingProgress(BytesProgressed&, CharDecrypted&)
Public Event DecryptingDone()
Public Event InterpretingBegin()
Public Event InterpretingDone()
Public Event InterpretingProgress(FasCmdlineObj As FasCommando)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long

Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long

Private Const RT_RCDATA = 10&

Private Const RT_STRING = 6&
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long


Dim FasCmdline As FasCommando

Dim Label_Push_col As New Collection
Dim LevelStack As New Stack
Dim level As Integer



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



FrmMain.ListView1.ListItems.Clear

Debug.Print vbCrLf & vbCrLf & vbCrLf & time & " " & String(60, "=")
Debug.Print "Interpreting FasResourceData..."
   
      ReDim FasStringtable(ResourceStreamVars)
      FasStringtables(0) = FasStringtable
      
      ReDim FasStringtable(FunctionStreamVars)
      FasStringtables(1) = FasStringtable
      
      FasStringtable = 0
      
      Dim ResourceStream As New StringReader
      FasResourceData.Position = 0
      ResourceStream.Data = FasResourceData.Data
      InterpretStream ResourceStream, FasResourceDataInterpeted
Debug.Print "Finish"


'next stringtable
'Set FasStringtable = FasStringtables(2)

Debug.Print String(60, "^")
Debug.Print "Interpreting FasFunctionData..."
      
      FasStringtable = 1
      
      Dim FunctionStream As New StringReader
      FasFunctionData.Position = 0
      FunctionStream.Data = FasFunctionData.Data
'      FasStack.ESP = 0
      InterpretStream FunctionStream, FasFunctionDataInterpeted
Debug.Print "Finish"



'Fill Functionlistwindow
frmFunction.Clear
Dim objtmp As FasFunction

For Each objtmp In FasFunctionnames
   frmFunction.Lb_Functions.AddItem objtmp.FuncName & "  " & objtmp.Startoffset
Next
frmFunction.Visible = True

'Fill Stringlistwindow
frmStrings.Clear
Dim objtmp1
For Each objtmp1 In FasStringtables(0)
   frmStrings.Lb_Strings.AddItem objtmp1 & "  0"
Next

frmStrings.Lb_Strings.AddItem String(50, "-")

For Each objtmp1 In FasStringtables(1)
   frmStrings.Lb_Strings.AddItem objtmp1 & "  1"
Next

frmStrings.Visible = True



Dim lsp_Filename
lsp_Filename = Filename & "_.lsp"

On Error Resume Next
Kill lsp_Filename

Open lsp_Filename For Output Shared As 2
Dim item
For Each item In FrmMain.LispFileData.Storage
   Print #2, item
Next
Close #2

'InterpretFile_err:
Close #1
'Err.Raise Err, , Err.Description
'Stop
'



End Sub
Private Sub InterpretStream(FileStream As StringReader, Target As Collection)
   RaiseEvent InterpretingBegin
   
      InterpretStream_rek FileStream, Target
   RaiseEvent InterpretingDone
End Sub
Private Sub InterpretStream_rek(FileStream As StringReader, Target As Collection, Optional StopAtOffset = &H7FFFFFFF, Optional level = 0, Optional Stackadd = 0, Optional ANDDest = 0, Optional InterpretStream_AND_6A_Helper = 0)

On Error GoTo InterpretStream_rek_err
   Dim tmp, dummy, ax As Long, cx As Long, ebx, edi
'   Dim case_43_counter  'helps to decide which stringtable to use
   Dim disasm As String
   
   Dim StartEsp&
   StartEsp = FasStack.ESP
   FasStack.ESP = FasStack.ESP + Stackadd
   

   With FileStream
      
      
      Do
         
       ' Create FasCmdline Object
         Set FasCmdline = New FasCommando
            
       ' Store current Offset
         FasCmdline.Position = .Position
         
       ' Initialise FasCmdline Object
         FasCmdline.Commando = .int8
         Set FasCmdline.Parameters = New Collection
         Dim param As Collection
         Set param = FasCmdline.Parameters
         
       ' add FasCmdline Object to Target collection
'         Target.Add FasCmdline
         FasCmdline.Interpreted = ""
       
       
         Label_Pop FasCmdline.Position
       
       ' Process Commands
         Select Case FasCmdline.Commando
         
'FunctionHead
         Case &H14:
            
'Debug.Assert FasStack.ESP <= 2

'FasStack.ESP = 0
            
            disasm = "defun main "
            On Error Resume Next
            disasm = "defun " & FasFunctionnames(CStr(.Position - 1))
            
            FasCmdline.Interpreted = "(" & disasm
            
            On Error GoTo InterpretStream_rek_err
            Dim var1, var2, var3, var4
            param.Add .int8 'Name of defun
            
            param.Add .int8   'esp=-.int8
            param.Add .int8
            

            param.Add .int8 'and         eax,0FFFE
            ax = param(4)
            cx = param(1)
          ' ax and cx represents a 15 bit value with the order axcx
          ' the last bit of ax is discarded
            var1 = ((ax And &HFE) * &H80) Or cx
'            Debug.Assert var1 <> 0
            
            var2 = param(2)
            var3 = param(3)
          
          ' only last bit matters
            var4 = CBool(param(4) And &H1)
            
            disasm = disasm & " with " & Hex(var1) & " vars, Arg_min: " & Hex(var2) & ", Arg_max: " & Hex(var3) & ", do_cleanup: " & Hex(var4)
            'pop
         
         Case &H15 'note: so far same as case 14
         
            disasm = "defun-q? main "
            On Error Resume Next
            disasm = "defun-q? " & FasFunctionnames(CStr(.Position - 1))
            
            FasCmdline.Interpreted = "(" & disasm
            
            On Error GoTo InterpretStream_rek_err
            param.Add .int8 'Name of defun
            
            param.Add .int8   'esp=-.int8
            param.Add .int8
            

            param.Add .int8 'and         eax,0FFFE
            ax = param(4)
            cx = param(1)
          ' ax and cx represents a 15 bit value with the order axcx
          ' the last bit of ax is discarded
            var1 = ((ax And &HFE) * &H80) Or cx
'            Debug.Assert var1 <> 0
            
            var2 = param(2)
            var3 = param(3)
          
          ' only last bit matters
            var4 = CBool(param(4) And &H1)
            
            disasm = disasm & " with " & Hex(var1) & " vars, Arg_min: " & Hex(var2) & ", Arg_max: " & Hex(var3) & ", do_cleanup: " & Hex(var4)
            'pop
         
         
         
         Case &H17
            'seems to be used in FAS2-File only
            
            disasm = "(FAS2) defun main "
            On Error Resume Next
            disasm = "(FAS2) defun " & FasFunctionnames(CStr(.Position - 1))
            On Error GoTo InterpretStream_rek_err
            
            
            param.Add .int16 '
            disasm = disasm & FasStringtables(FasStringtable)(param(1))
            'push clng(3)
         
'Const nil & T
         Case &H1
            disasm = "Push nil "
            FasStack.push "nil"
         Case &H2
            FasStack.push "T"
            disasm = "Push T" 'command AND?
'Argument=param
         Case &H5:
            param.Add .int8
            tmp = "FuncArg[" & param(1) & "]"
            disasm = "push(FSL) " & tmp
            FasStack.push tmp

'FuncArgument=last
         Case &H8: 'FSL
            param.Add .int8
            disasm = "FuncArg[" & param(1) & "] = " & FasStack.pop '" [value above]"
            FasCmdline.Interpreted = disasm
            'disasm = "setq_arg " & FasStringtables(FasStringtable)(param(1)) & " = " & FasStack.pop '" [value above]"
'            disasm = "losid::mark(void) mov[[ebp-18h]+[param1]*4], pop " & tmp 'FasStringtables(FasStringtable)(param(1))
'Argument=0
         Case &H64: 'FSL
            param.Add .int8
            disasm = "FuncArg[" & param(1) & "] = 0"
            'disasm = "arg " & param(1) & "[" & FasStringtables(FasStringtable)(param(1)) & "] = 0"
         
         Case &H5C
            param.Add .int16
            disasm = "Push FuncArg[" & param(1) & "]"
            FasStack.push "FuncArg[" & param(1) & "]"  'this is not correct, but may work for now"
                              'don't forget to adapt 'case 43' also if you modify this
'           'FuncArg=[ebp-18h]
         Case &H5D
            param.Add .int16
            tmp = FasStack.pop
            disasm = "FuncArg[" & param(1) & "] = " & tmp '[module]
'           'FuncArg=[ebp-18h]
         Case &H5E
            '100%
            param.Add .int16
            disasm = "clear FuncArg[" & param(1) & "]" '  mov [ebp-18h] or [local.6] <= Funcarg
         

'BinValues ops
         Case &H3
           param.Add .int16  'String?
           tmp = FasStringtables(FasStringtable)(param(1))
           disasm = "Push value of [" & tmp & "]"
           FasStack.push tmp
         Case &H4
            'goes together with 2e
             param.Add .int8
             param.Add .int8
             tmp = "stream " & param(2) & " [" & param(1) & "]"
             disasm = "push " & tmp 'lnode*
             'Debug.Assert param(1) = 1 'listselector or rewinder - seems to be always 1
             'Debug.Assert (param(2) = 0) Or (param(2) = 1) Or (param(2) = 2) Or (param(2) = 4)
             
             FasStack.push tmp
         Case &H6:
            param.Add .int16
            disasm = "setq " & FasStringtables(FasStringtable)(param(1)) & " " & FasStack.pop '" [value above]"
            If FasStack.ESP > (StartEsp) Then
               FasStack.pop
               FasStack.push "(" & disasm & ")"
            End If
            FasCmdline.Interpreted = "; (" & disasm & ")"

         Case &H1A
            param.Add .int16
            disasm = "FSL_func_setq " & FasStringtables(FasStringtable)(param(1)) & " " & FasStack.pop
            
            FasCmdline.Interpreted = "(" & disasm & ")"
            
         Case &H1B
            param.Add .int16
            disasm = "FSL_setq " & FasStringtables(FasStringtable)(param(1)) & " " & FasStack.pop
         
            FasCmdline.Interpreted = "(" & disasm & ")"
            
         Case &H7:
            param.Add .int8 'List(>index<)
            param.Add .int8 'List(>index<)
            disasm = "Copy Element " & param(1) & " to " & param(2) & " ListBase=[61A49790]"
            '61A49790
  
         Case &H10
            param.Add .int8 'steps forward or backward in a list ?
            param.Add .int8
            disasm = "Push??? FuncArg " & param(1) & ", " & param(2)
            FasStack.push disasm
  
         Case &H9:
            param.Add .int16Sig
            tmp = FasStringtables(FasStringtable)(param(1))
            disasm = "push " & tmp
            FasStack.push tmp
         Case &HA
            disasm = "pop dummy (decrease stack)" & vbCrLf  ' "sub-4"
            FasCmdline.Interpreted = FasStack.pop
         Case &HB:
            FasStack.push FasStack.PreviewPop
            disasm = "Push " & FasStack.PreviewPop & "[Last element again]"
              
         Case &HC
            param.Add .int16Sig
            FasStack.push FasStringtables(FasStringtable)(param(1))
            disasm = "FSL_Push " & FasStack.PreviewPop
            
            
         Case &HD, &H3C ' identical code in vl.arx (not point same points in cmd-switch table)
         '0d->0f_jmp only used in fsl
         '0f->3C_with_jump_backward
            param.Add .int16Sig 'old
            disasm = "if (" & FasStack.pop & "==0) then jmp to " & param(1) + .Position
            FasStack.push "'Retval' of IF"
            
            FasCmdline.Disassembled = Space(3 * level) & disasm
            RaiseEvent InterpretingProgress(FasCmdline)
            
            If param(1) > 0 Then
               InterpretStream_rek FileStream, Target, .Position + param(1), level + 1, -1
'               If .int8 = &H1A Then 'is next cmd setq?
'                  qw
'                  param.Add .int16
'                  disasm = "FSL_setq " & FasStringtables(FasStringtable)(param(1)) & " " & FasStack.pop
'
'               Else
'                  .Move -1
                 'FasStack.pop
                 GoTo noOutput
'               End If
               
            Else
               FasStack.pop
               GoTo noOutput
            End If

         
         Case &HE, &H3D ' identical code in vl.arx (not point same points in cmd-switch table)
            param.Add .int16Sig
            
            disasm = "if (" & FasStack.PreviewPop & ") then pop else jmp to " & param(1) + .Position
            
            FasCmdline.Disassembled = Space(3 * level) & disasm
            RaiseEvent InterpretingProgress(FasCmdline)
            
            InterpretStream_rek FileStream, Target, .Position + param(1), level + 1, -1
        '    If level = 0 Then FasStack.push "'Retval' of IF"
        
'            qw
            GoTo noOutput
         
         Case &H67: 'normal if then else 57 used for
         '( while testexpr [expr])
         '( repeat int [expr])
         '( foreach name lst [expr]) and
         '( if testexpr thenexpr elseexpr)
            
            param.Add .int32
            tmp = FasStack.pop
         
           ' is it Repeat?
            Dim bRepeat As Boolean
            bRepeat = Last_FasCmdline.Commando = &H4B '(<= a b)
            
            If bRepeat Then  '  push is (0 <= FuncArg[0])
               disasm = "(Repeat " & tmp & " ... else goto " & .Position + param(1) ' If [edi]==0
               tmp = "(Repeat " & tmp
            Else
               disasm = "If (" & tmp & ")... else goto " & .Position + param(1) ' If [edi]==0
               tmp = "(if " & tmp
            End If
            
            
            'FasCmdline.Interpreted = tmp
            'FasStack.pop
            'FasStack.push tmp
            
            FasCmdline.Disassembled = Space(3 * level) & disasm
            RaiseEvent InterpretingProgress(FasCmdline)
            
            'process  'Then' Branch
            InterpretStream_rek FileStream, Target, .Position + param(1), level + 1
            
          ' is it normal_If?
            If Last_FasCmdline.Commando = &H57 Then 'Note: &H57 is Goto xxx
             ' 'else' branch of if
               If Last_FasCmdline.Parameters(1) >= &H0 Then
                
                  InterpretStream_rek FileStream, Target, Last_FasCmdline.Position + Last_FasCmdline.Parameters(1), level + 1
                  'FasStack.pop
                  FasStack.Current = "Then OR Else"
               Else
                ' is it While  (For Each)?
                  If bRepeat Then
                     FasStack.Current = "Repeat"
                  Else
                     Dim nextcmd&
                     nextcmd = .int8
                     .Move -1
                     If nextcmd = &H19 Then 'Note: 19    clear args+vars
                        FasStack.Current = "For Each"
                     Else
                        FasStack.Current = "While"
                     End If
                     
                  End If
               End If
            Else
               Err.Raise vbObjectError, "", "Then-block of normal_if(&h67) did end with goto(&H57)"
               Stop
            End If
            
        
'            FasCmdline.Interpreted = FasStack.pop & ")"
 '           disasm = Space(3 * level) & "END IF"
            GoTo noOutput
            
         Case &H69
            'seldom use
            param.Add .int32
            'pop condition always
            disasm = "If (" & FasStack.pop & ") ... Else Goto " & .Position + param(1)
         
         Case &H68 'Cond
            param.Add .int32
            tmp = .Position + param(1)
            disasm = "Cond If (" & FasStack.PreviewPop & ") Goto " & tmp & " Else .pop"
            
'            Label_Push CStr(tmp), "cond"
            
            FasCmdline.Disassembled = Space(3 * level) & disasm
            FasCmdline.Interpreted = "(cond " & FasStack.PreviewPop & " ("
            RaiseEvent InterpretingProgress(FasCmdline)

            InterpretStream_rek FileStream, Target, .Position + param(1), level + 1, -1
                      ' is it normal_If?
            If Last_FasCmdline.Commando = &H57 Then
               FasCmdline.Interpreted = "normal cond"
            Else
               FasCmdline.Interpreted = "it's OR skip next 6 bytes -> " & .Position + 6
            End If
        
            GoTo noOutput

         Case &H6A: 'AND
            param.Add .int32
            tmp = .Position + param(1)
            disasm = "and_If (" & FasStack.PreviewPop & ") .pop Else goto " & tmp
'            FasCmdline.Disassembled = Space(3 * level) & disasm
'            RaiseEvent InterpretingProgress(FasCmdline)

            Label_Push CStr(tmp), "and"
'            If ANDDest = tmp Then
'              'Next/inner andif
'               InterpretStream_rek FileStream, Target, .Position + param(1), level + 1, -1, tmp, InterpretStream_AND_6A_Helper
'               InterpretStream_AND_6A_Helper = InterpretStream_AND_6A_Helper + 1
'            Else
'              'First/ outer andif
'               InterpretStream_AND_6A_Helper = 1
'               InterpretStream_rek FileStream, Target, .Position + param(1), level + 1, -1, tmp, InterpretStream_AND_6A_Helper
'             '  FasStack.ESP = FasStack.ESP + InterpretStream_AND_6A_Helper - 1
'             '  FasStack.push "(and " & Join(FasStack.popArray(InterpretStream_AND_6A_Helper)) & ")"
'               InterpretStream_AND_6A_Helper = 1
''               disasm = Space(3 * level) & "-------"
'            End If
'            GoTo noOutput
            
         Case &HF 'fsl - together with 0d_if
            param.Add .int16
            disasm = "jump to " & param(1) + .Position
            FasCmdline.Disassembled = Space(3 * level) & disasm
            RaiseEvent InterpretingProgress(FasCmdline)
            
            If param(1) > 0 Then
               InterpretStream_rek FileStream, Target, .Position + param(1), level + 1
            End If
        
            'GoTo noOutput
            
            
         Case &H57
            param.Add .int32
            disasm = "goto " & param(1) + .Position
            'disasm = "jump [far] over " & param(1) & " bytes to " & param(1) + .Position
'            If (param(1) > 0) And (level > 0) Then
'               disasm = "ELSE " & disasm
'               StopAtOffset = StopAtOffset + param(1)
'            End If

            
         Case &H21
            param.Add .int16
            disasm = "end defun ;cleanup " & param(1) & " vars (" & param(1) & " " & Join(FasStack.popArray(param(1) + 2), ",") & ")"
            'select case param1 '[61A49788]
            'case 1..10    5,6 invalid cmd  2 exit
'         Case &H22
'            Stop
         Case &H3E

            disasm = "end Defun if " & FasStack.pop & " != ZERO"
         Case &H3F
            disasm = "end Defun if " & FasStack.pop & " == ZERO"
            
         Case &H16:
            disasm = "end Defun" & vbCrLf ' Retval: " & FasStack.PreviewPop
            FasCmdline.Interpreted = FasStack.pop ' & " )" 'new
            
            FasStack.ESP = 0
         
         Case &H1C
            disasm = "init done " & FasStack.pop
            FasStack.ESP = 0
            'ebp+c=0
            '[61A49790]=[[61A49790]+C]
            'pop, reset varbasepionter and something else...
 '           Set FasStringtable = FasStringtables(1)
         Case &H34
            param.Add .int8 ' Func Arguments
            param.Add .int8 'Flags
            FasStack.push FasStack.pop & " " & Join(FasStack.popArray(param(1)))
            disasm = FasStack.PreviewPop & ", " & _
                                         CBool(param(2) And 1) & ", " & _
                                         CBool(param(2) \ 4 And 1) & ", " & _
                                         CBool(param(2) \ 2 And 1) & ", "
            'if
            
            

         Case &H35: param.Add .int8  'Params are above (stored in [ebp+c])
                    param.Add .int16 'String
                    param.Add .int8 'used bits(0),bits(2),bits(1)
                    Debug.Assert (param(3) = 3) Or (param(3) = 1)
                    'Debug.Assert .Position <> 2735
                    disasm = IIf(param(3) = 1, "Internal ACAD Command ", "") & _
                             FasStringtables(FasStringtable)(param(2)) & " " & _
                               param(1) & " Params are above..."
                    
                    FasStack.push "(" & FasStringtables(FasStringtable)(param(2)) & " " & _
                                             Join(FasStack.popArray(param(1))) & ")"
'Convert to funcptr
         Case &H51
              
              param.Add .int8 'gets overwrite with 45 or 5A
              param.Add .int16  'String     ->long(lo)
              param.Add .int16  'test if 00 ->long(hi)
              '"<- Convert to funcptr and ebx--"
              tmp = FasStringtables(FasStringtable)(param(2))
              Select Case tmp
               Case "string-resource"
                  Dim tmpstr1$, tmpstr2$, tmpstr3$
                  tmpstr1 = FasStack.pop
                  tmpstr2 = FasStack.pop
                  tmpstr3 = """" & _
                           LoadResStringDll(tmpstr1, _
                                             App.Path & "\" & Split(tmpstr2, """")(1) _
                           ) & """"
                  disasm = tmpstr3 & " " & tmp & " from " & tmpstr2
                  FasStack.push tmpstr3 '& "<-" & tmpstr2
               Case Else
                  disasm = tmp & " " & param(1) & " Params are above <- Convert to funcptr"
                  'FasStack.ESP = FasStack.ESP - param(1)
                  'FasStack.push "Retval of " & FasStringtables(FasStringtable)(param(2))
                  tmpstr2 = Join(FasStack.popArray(param(1)))
                  
                  tmpstr1 = FrmMain.LispFileData.pop
                  
                  If InStr(tmpstr1, ")") Then
                    'Defun xyz (x1, x2 / a1, a2)
                     tmpstr1 = Replace(tmpstr1, ")", Replace(tmpstr2, "(", " / "))
                  Else
                    'Defun xyz (/ a1, a2)
                     tmpstr1 = tmpstr1 + Replace(tmpstr2, "(", "(/ ")
                  End If
                  
                  If tmp = "_al-bind-alist" Then
                     FasStack.push tmpstr1
                  Else
                     FasStack.push "(" & FasStringtables(FasStringtable)(param(2)) & " " & _
                          tmpstr2 & ")"

                  End If


              End Select
            
         Case &H45 'Function in Ram (Converted)
            param.Add .int8
            param.Add .int32
            disasm = "Memdumpfunc !!! Call " & param(2) & " (" & Join(FasStack.popArray(param(1))) & ")"
            FasStack.push disasm
         
         Case &H5A  'Function in Ram (Converted)
            param.Add .int8 'Number of Stack elements; 0..3 valid else "error in XNCALL"
            param.Add .int32 '->fctptr "..\KERN\FNS\XRUN.CPP"

            disasm = "Memdumpfunc !!! Call VL.ARX!" & param(2) & " (" & Join(FasStack.popArray(param(1))) & ")"
            FasStack.push disasm
         
         
         
         Case &H5F
            param.Add .int8 '0
            param.Add .int32 '1
            
            disasm = ""
            On Error Resume Next
            disasm = FasFunctionnames(CStr(param(2)))
            On Error GoTo InterpretStream_rek_err
 
            disasm = "jmp1 " & disasm & " at " & param(2) & "(" & Join(FasStack.popArray(param(1)), " ") & ")"
            FasStack.push "Retval of " & disasm
            
         
         
         Case &H53
         '<vl.public: virtual short closure::ltid(void)>, vl.closure::ltid
               disasm = "53! " & Join(FasStack.popArray(2))
               FasStack.push "Retval of " & disasm
             ' disasm = "Reset? " & FasStack.pop & " & " & FasStack.pop
              '[dgcpro::head]
              '<vl.public: virtual short cons::ltid(void)>), vl.cons::ltid

              '"bad argument list tail:"
         
         Case &H54
              'disasm = "unknow delta F"
              disasm = "54! " & Join(FasStack.popArray(2))
              FasStack.push "Retval of " & disasm
         
         Case &H2E
            param.Add .int8 'flag ?
            disasm = "StackCallAndJmp " & FasStack.pop & " (" & Join(FasStack.popArray(param(1)), ",") & ")"
         
         Case &H2F
            param.Add .int8
            param.Add .int16
            disasm = "CallAndJmp " & FasStringtables(FasStringtable)(param(2)) & " (" & Join(FasStack.popArray(param(1)), ",") & ")"
            
         Case &H60
            
            param.Add .int8 '0
            param.Add .int32 '1
            
            disasm = ""
            On Error Resume Next
            disasm = FasFunctionnames(CStr(param(2)))
            On Error GoTo InterpretStream_rek_err
 
            disasm = "jmp2_nopop " & disasm & " at " & param(2) & "(" & Join(FasStack.popArray(param(1)), " ") & ")"
            FasStack.push "Retval of " & disasm

         Case &H61
            param.Add .int8
            param.Add .int32
            disasm = "[" & param(2) & "]"
            On Error Resume Next
            disasm = FasFunctionnames(CStr(param(2)))
            On Error GoTo InterpretStream_rek_err
 
            disasm = "continue at " & disasm & " " & Join(FasStack.popArray(param(1)), " ") & ")"
            FasStack.push "(" & disasm & ")"
            
         Case &H18:
            param.Add .int16 '
            tmp = FasStack.popArray((param(1) \ 2) * 2)
            disasm = "Init ("
            If IsArray(tmp) Then
            
               For i = LBound(tmp) To UBound(tmp) Step 1
                  If i Mod 2 Then
                     disasm = disasm & tmp(i)
                     FasCmdline.Interpreted = FasCmdline.Interpreted & "  " & tmp(i)
                  Else
                     disasm = disasm & "=" & tmp(i) & "  "
                  End If
               Next
               
            End If
            
            tmpstr1 = FrmMain.LispFileData.pop
            tmpstr2 = "(" & LTrim$(FasCmdline.Interpreted) & ")"
            If InStr(tmpstr1, "defun") Then
               FasCmdline.Interpreted = tmpstr1 + " " + tmpstr2
            Else
               FasCmdline.Interpreted = tmpstr2
            End If
            
            

            tmp = Empty
            disasm = disasm & ") " & ((param(1)) And 1)
'Debug.Assert (param(1) And 1) = 0
            
         Case &H19:
            param.Add .int16
'            Debug.Assert (param(1) And 1) = 1  'stop when this is 0 - for further exmiation?
            'tmp = param(1) \ 2 + 1
            disasm = "clear " & param(1) & " args+vars" '& _
               IIf((param(1) And 1), "", " last func in stream !")
'            FAsStack.ESP = FAsStack.ESP - tmp
                           
'         Case &H1D
'            Stop
         Case &H1E
            param.Add .int8
            disasm = "??? alpha " & FasStack.pop
         Case &H1F
            param.Add .int8
            disasm = "??? beta"
'Symbol-Handling Functions
         Case &H23 ' Null and Not is same - Typically
                   'null' is used for lists, and
                   'not' is used for other data types along with some types of control functions.
               disasm = "Is " & FasStack.PreviewPop & " == 0 ?"
               FasStack.push "(NULL " & FasStack.pop & ")"

         Case &H24
            disasm = "Is " & FasStack.PreviewPop & " no list?"
            FasStack.push "(atom " & FasStack.pop & ")"
         Case &H25
            param.Add .int8
            tmp = "FuncArg[" & param(1) & "]"
            disasm = "IsSym (FSL) " & tmp & " & " & FasStack.PreviewPop 'ltid
'            FasStack.push tmp
            
'            disasm = "push (cmp " & param(1) & ", " & FasStack.pop
         Case &H26
            tmp = FasStack.pop & " - " & FasStack.pop
            FasStack.push tmp
            disasm = "push (" & tmp & ")"
            
'Listmanipulations
         Case &H28 'Get first element
            disasm = "push list element from (" & FasStack.PreviewPop & ")"
            FasStack.push "(car " & FasStack.pop & ")"
            'tmp = "list element from (" & FasStack.pop & ")" 'losid::mark(void)'[ebx+08h]
            'disasm = "push " & tmp
            'FasStack.push tmp
         Case &H29  'Get first elements
            disasm = "push next elements from (" & FasStack.PreviewPop & ")"
            FasStack.push "(cdr " & FasStack.pop & ")"
            
'            tmp = "next list element from (" & FasStack.pop & ")" 'losid::mark(void) '[ebx+0Ch]
'            disasm = "push " & tmp
'            FasStack.push tmp
         Case &H2A  'Insert element at beginning
'            If FasStack.ESP > 1 Then
               tmp = FasStack.popArray(2)
'            Else
'               tmp = "Stackerror!"
'            End If
'            'disasm = "defFunc " & Join(tmp, " .. ")
            
            disasm = "(Insert " & Join(tmp, " at beginning of ")
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
            FasStack.push "(cons " & Join(tmp, " ") & ")"
         Case &H2C
            param.Add .int8 'index
            tmp = FasStack.pop & " [" & param(1) & "]" '& FasStringtables(FasStringtable)(param(1))  '"losid_subtype_p"
            FasStack.push (tmp)
            disasm = tmp
            
         Case &H2D
            param.Add .int8 'index
            tmp = FasStack.pop & " = " & FasStack.pop & " [" & param(1) & "]" '& FasStringtables(FasStringtable)(param(1)) 'losid_subtype_p
            FasStack.push tmp
            disasm = tmp
            
         Case &H32:

            param.Add .int8
          ' convert unsigned char to signed char
            tmp = param(1)
            If tmp > &H7F Then
               tmp = (tmp And &H7F) Or &HFF80
            End If
            
           'push .int32
            disasm = "push  " & Format(tmp, "00")
            FasStack.push tmp

         Case &H33:
            param.Add .int32 'var1=v+v
            'push .int32
            disasm = "push " & Format(param(1), "00000000")
            FasStack.push param(1)
         
         Case &H55
'            param.Add .int16sig 'Number of Strings follow
           
           'Read all Strings
            For i = 1 To .int16 'param(1)
'               param.Add .int16 'Stringlength
                                                                            '=Last Parameter

               tmp = Replace(.FixedString(.int16), vbLf, "\n")
'               tmp = .fixedString(param(param.Count))
               param.Add tmp
             
             ' Add String to StringTable
               FasStack.push """" & tmp & """"

               
            Next
            disasm = "String [lnode]"
            
         Case &H56, &H5B 'Create Stringtable
            'FSL 56
            disasm = "load Commands"
            
            Do
             ' get String
               tmp = .zeroString
             
             ' is the string the "Terminatorstring" (=Zerolength)
               If Len(tmp) = 0 Then Exit Do
             
             ' store String
               param.Add tmp
               
             ' Add String to StringTable
               FasStack.push tmp
            Loop While True
         
         
         Case &H40
            param.Add .int16 '+0
            param.Add .int32 '+2
            disasm = "??? memMov sth."
            
            FasStack.push disasm
         
         Case &H37
            param.Add .int16Sig
            tmp = "ListObject B with " & param(1) & " elements"
            disasm = "push " & tmp '[cons]
            'pop * number of param(1); push FuncListObject
'            FAsStack.ESP = FAsStack.ESP - param(1)
            FasStack.push "<list( " & Join(FasStack.popArray(param(1)), ",") & ")>"
            
            
         Case &H38:
               disasm = "Convert last element on stack"
               
         Case &H3B
              param.Add .zeroString
              FasStack.push param(1)
              disasm = "[real] load numberString"
         
         Case &H39   'Choose Array for store strings
            param.Add .int16
            tmp = "ListObject with " & param(1) & " elements"
            disasm = "push " & tmp '[cons]
            'pop * number of param(1); push FuncListObject
'            FAsStack.ESP = FAsStack.ESP - param(1)
            FasStack.push "(" & Join(FasStack.popArray(param(1))) & ")"
            
         Case &H3A
            tmp = FasStack.pop
            var1 = FasStack.pop
            var2 = FasStack.pop
           
            Dim FasFunction As FasFunction
            Set FasFunction = New FasFunction
            
            FasFunction.Startoffset = var1
'            FasFunction.Endoffset = ???
            
' DIRTY-Code: the moduleparameter (var2) is not used - so sometimes we have to ".remove" before ".add"
            On Error Resume Next
            FasFunctionnames.Remove CStr(FasFunction.Startoffset)
            On Error GoTo InterpretStream_rek_err
            
            FasFunctionnames.Add FasFunction, CStr(FasFunction.Startoffset)
            
            disasm = "defFuncName " & tmp & ", StartOffset " & var1 & " Modul " & var2 & vbCrLf '[xsubr]
            FasStack.push "<Func> " & tmp
            FasFunctionnames.item(CStr(var1)) = tmp
         Case &H59
            disasm = "Execute or load sth"
            
         Case &H43
            param.Add .int16
            param.Add .int16
            tmp = FasStack.pop
            disasm = "mov " & param(2) & " Stackelements to VarPos " & param(1) & " at " & IIf(tmp = "nil", "this stream", "FuncStream")
                                          
            'case_43_counter = case_43_counter + 1
            tmp = IIf(tmp = "nil", 0, 1) 'Choose(case_43_counter, 0, 1, 0)
            'var1=FAsStack.popArray(
            For i = param(2) - 1 + param(1) To param(1) Step -1
               FasStringtables(tmp)(i) = FasStack.pop
            Next

            FasStack.push "FasStringtables " & tmp
'           fas43count = fas43count + 1 '<- dirty dirty smuddcode !!!
           'Create new stringtable only if nessary
'            If (fas43count = 1) Or (fas43count = 3) Then
'               FasStringtables.Add FasStringtable
'               Set FasStringtable = New Collection
'            End If


'Binary fixnum operations
        
         Case &H46
            disasm = "+ " & Join(FasStack.popArray(2))
            FasStack.push "(" & disasm & ")"
         
         Case &H47
            disasm = "- " & Join(FasStack.popArray(2))
            FasStack.push "(" & disasm & ")"
            
         Case &H48
            disasm = "* " & Join(FasStack.popArray(2))
            FasStack.push "(" & disasm & ")"
         
         Case &H49
            disasm = "/ " & Join(FasStack.popArray(2))
            FasStack.push "(" & disasm & ")"
            
         Case &H4A
            disasm = "mod " & Join(FasStack.popArray(2))
            FasStack.push "(" & disasm & ")"

         Case &H4B
            disasm = "<= " & Join(FasStack.popArray(2))
            FasStack.push "(" & disasm & ")"
         
         Case &H4C
            disasm = ">= " & Join(FasStack.popArray(2))
            FasStack.push "(" & disasm & ")"
         
         Case &H4D
            disasm = "< " & Join(FasStack.popArray(2))
            FasStack.push "(" & disasm & ")"
         
         Case &H4E
            disasm = "> " & Join(FasStack.popArray(2))
            FasStack.push "(" & disasm & ")"
         
'unary fixnum operations
         Case &H4F
            tmp = FasStack.pop
            FasStack.push "(1+ " & tmp & ")"
            disasm = "inc " & tmp
         Case &H50
            tmp = FasStack.pop
            FasStack.push "(1- " & tmp & ")"
            disasm = "dec " & tmp
         
         
         Case &H65
            disasm = "TraceIn? " & FasStack.pop
         Case &H66
            '100%
            disasm = "TraceOut? [61A4974C]=[61A4974C]-4"
         
         
         Case &H20, &H62, &H63
              disasm = "nop"
            
         Case &H11, &H12, &H13, &H1D, &H22, &H27, &H2B, _
              &H30, &H31, &H36, &H41, &H42, &H44, &H52, &H58
            disasm = "Stop Processing this stream"
            
         Case Else
            disasm = "Invalid commando"
'            If vbYes = MsgBox("Invalid fas commando at " & .Position & vbCrLf & "Stop precessing ?", vbExclamation + vbYesNo + vbDefaultButton2) Then Stop
'            Stop 'due to invalid commando
         End Select
         
         
'         FasCmdline.Interpreted = Interp
         FasCmdline.Disassembled = Space(3 * level) & disasm
         RaiseEvent InterpretingProgress(FasCmdline)
         If FrmMain.Chk_Cancel Then
            Err.Raise ERR_GUI_CANCEL, "", tmp
         End If

         
         
noOutput:
         
         Set Last_FasCmdline = FasCmdline
         
      Loop Until .EOS Or (.Position >= StopAtOffset)
   
'   If FasStack.ESP > StartEsp Then
'      FasCmdline.Interpreted = FasStack.pop
'      FasCmdline.Disassembled = Space(3 * level) & "Retval"
'      RaiseEvent InterpretingProgress(FasCmdline)
'   End If

   FasStack.ESP = StartEsp

End With

InterpretStream_rek_err:
Select Case Err
   Case 0
   Case 1001 + vbObjectError 'pop not possible
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
      Resume Next
      Err.Raise Err, , Err.Description
End Select


End Sub

Private Sub getStreamData(Datalength, INPUT_NumberOfVars, INPUT_NewStream As FileStream)
   Dim StreamTerminatorChar
   Dim tmpchar$
   With File

    
   ' Get FunctionStreamVars
     tmpchar = SkipWhiteSpaceEx
     If IsNumeric(tmpchar) = False Then Err.Raise vbObjectError, , "Invalid File Format - Could not get Number of StreamVars."
     
     Do While IsNumeric(tmpchar) And (.EOS = False)
           INPUT_NumberOfVars = INPUT_NumberOfVars & tmpchar
           tmpchar = .char
     Loop
     
   ' Get StreamBeginChar (and store it in 'StreamTerminatorChar')
     tmpchar = Skipwhitespace
     If tmpchar = "!" Then
        StreamTerminatorChar = .char
     Else
        StreamTerminatorChar = tmpchar
     End If
  
  ' Store start of code
    CodeStart = .Position
    
  ' Fill INPUT_NewStream with data
    INPUT_NewStream.FixedString(-1) = .FixedString(Datalength)
    INPUT_NewStream.Position = 0
 

  ' Get next Char - if Char=StreamBeginChar the Stream is uncrypted _
                  - if not it's the length of the following key
    tmpchar = .char
    
  ' Set FileEncrypted flag
    Dim IsEncrypted As Boolean
    IsEncrypted = (tmpchar <> StreamTerminatorChar)

    If IsEncrypted Then
       
     ' Set & Check Keylength
       Dim keylength&
       keylength& = Asc(tmpchar)
       If keylength >= &H80 Then Err.Raise vbObjectError, , "Keylength is bigger than 128 Byte! Processing canceled."
       
       KeyStart = .Position
      
      ' Create FasKey & Fill it with data
        FasKey.Create .Filename & ".key", True, CreateFilesTemporary
        FasKey.FixedString(-1) = .FixedString(keylength)
        
        tmpchar = .char

    End If
    If tmpchar <> StreamTerminatorChar Then Err.Raise vbObjectError, , "StreamBeginChar and StreamTerminatorChar(after key) are not equal! Processing canceled."
 
   End With
 
    If IsEncrypted Then
            
          Dim KeyOld As Long, KeyNew As Long, value As Long
           
          FasKey.Position = 0
          KeyOld = FasKey.int8
         
          If KeyOld <> 0 Then
             
             RaiseEvent DecryptingBegin(INPUT_NewStream.Length)
             
             With INPUT_NewStream
           
              ' Create Filestream to write uncrypted data in FasResourceData
                Dim FasResourceDataWrite As FileStream
                Set FasResourceDataWrite = New FileStream
                FasResourceDataWrite.Create (.Filename)
              
              
              ' Set filepointer at begin of crypted data
                File.Position = CodeStart
            
                .Position = 0
                Do Until .EOS
              
                   If FasKey.EOS Then FasKey.Position = 0
                   If (Second(time) Mod 2) = 0 Then
'                     DoEvents
                   End If
                   
                   KeyNew = FasKey.int8
            
                   value = .int8 Xor KeyNew Xor KeyOld
                   
                   RaiseEvent DecryptingProgress(.Position, value)
                   
                   File.int8 = value
                   
                 ' write uncrypted data to FasResourceData
                   FasResourceDataWrite.int8 = value
                   
                   KeyOld = KeyNew
                   
'                   DoEvents
                Loop
              
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
               Else
                 
                 ' Set StreamTerminatorChar so Lisp will recognize this stream as uncrypted
                   File.Position = KeyStart - 1
                   File.char = StreamTerminatorChar
   
   ' ' optionally: Let the key start with \0 so the encrypting Loop will do nothing
   '                File.Position = KeyStart
   '                File.int8 = 0
               End If
                
               RaiseEvent DecryptingDone
                
            '     IsDecryptingDone = True
                
            '     FasResourceData.position = 0
            '     FasResourceData.fixedString(-1) = File.fixedString
                
            End With
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
           OutStream.Create .Filename & Ext, True, CreateFilesTemporary
         '...& Fill it with data
           getStreamData StreamLength, StreamVars, OutStream

         ' Store for use later in Hexeditor jump
           Offset_Start = CodeStart
           
           .getTerminatedString "#"         ' Skip whitespaces


   End With
End Sub
Private Sub FASStreamLoad( _
                           OutStream As FileStream, _
                           Ext As String, _
                           StreamLength&, _
                           StreamVars&, _
                           Offset_Start&)
   With File
   
         '1. Get FunctionStreamLength
           Dim tmpchar$
           tmpchar = SkipWhiteSpaceEx
           If IsNumeric(tmpchar) = False Then Err.Raise vbObjectError, , "Invalid File Format - Could not get Length of FasFunctionStream."
           
           Do While IsNumeric(tmpchar) And (.EOS = False)
                 StreamLength = StreamLength & tmpchar
                 tmpchar = .char
           Loop
         
         '2. Create FasFunctionDataStream...
           OutStream.Create .Filename & Ext, True, CreateFilesTemporary
         '...& Fill it with data
           getStreamData StreamLength, StreamVars, OutStream


         ' Store for use later in Hexeditor jump
           Offset_Start = CodeStart


   End With
End Sub



Public Function Create(Filename As String) As Boolean
    
   CreateFilesTemporary = FrmMain.Chk_cleanup.value = vbChecked
    
With File
  ' Store Filename
    mvarFileName = Filename
    
    .Create Filename
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
    
    Dim tmpchar$
    tmpchar = SkipWhiteSpaceEx
    
   'Is it a FAS or FSL File?
    If WHITESPACETABLE(Asc(tmpchar)) = WS_DASH Then
   
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
  
       Else
          Err.Raise vbObjectError + 1, "Create", "Invalid FSL_FILE_SIGNATURE. No valid Fsl File"
       End If
    
    Else
     'FAS_FILE INIT
 ' 0. Signature
       Dim FileSig As New clsStrCat
     ' Read all following Alphanumeric chars  (until offset 1024 - to speedup checking invalid files)
       Do While (WHITESPACETABLE(Asc(tmpchar)) = WS_ALPHANUMMERIC) And (.Position < 1024) 'note: tmpchar was initialed above this If-Block
         
         FileSig.Concat tmpchar
         tmpchar = File.char
       Loop
     
     ' rewind by one byte (because of 'tmpchar = File.char')
       File.Move -1
       
       Select Case FileSig
       Case FAS4_FILE_SIGNATURE, _
            FAS3_FILE_SIGNATURE, _
            FAS2_FILE_SIGNATURE, _
            FAS__FILE_SIGNATURE
         
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

       Case Else
         Err.Raise vbObjectError + 1, "Create", "Invalid Fas File. FAS__FILE_SIGNATURE not found."
       End Select
    End If
   
   .CloseFile

  End With

  RaiseEvent InitDone

  If Not (FrmMain.chk_Decryptonly.value = vbChecked) Then

   ' Start Interpreting
     Me.InterpretFile
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

Public Sub Label_Push(Label$, Command$)
   Dim adder As clsStrCat
   
   On Error Resume Next
   Set adder = Label_Push_col.item(Label)
   If Err Then
      
      LevelStack.push level
      
      Set adder = New clsStrCat
      adder.Concat Command
      adder.Concat " " & FasStack.pop
      Label_Push_col.Add adder, Label
   Else
      adder.Concat " " & FasStack.pop
   End If
   
   level = level + 1
   
    '& "(" & FasStack.PreviewPop & ")
End Sub

Public Sub Label_Pop(Label$)
   Dim adder As clsStrCat
   
   On Error Resume Next
   Set adder = Label_Push_col.item(Label)
   If Err = 0 Then
      Dim tmp
      tmp = "(" & adder.value & ")"
      FasStack.Current = tmp
'      FasCmdline.Interpreted = tmp
      level = LevelStack.pop
      
   End If
End Sub
