VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Brancher_Cond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PartsSeperatorOnTheLine As String = " "
Private Const PartsSeperator As String = PartsSeperatorOnTheLine & vbCrLf

Private Brancher As Fas_Brancher
Private FileStream As FileStream
Private FasStack As Stack

Private Const Goto_0x57_Size As Long = 1 + 4 'Size for 0x57 Goto opCode

Private cond_items As New Collection
Private And_Items As New Collection

Private cond_body_end

Public Sub create(Fas_Brancher As Fas_Brancher, FileStream_ As FileStream)
   Set Brancher = Fas_Brancher
   Set FileStream = FileStream_
   
   Set FasStack = Brancher.FasStack
End Sub

'(cond [((test) (result ...)) ...])
Public Sub HandleCond(FasCmdline As FasCommando) ', cond_item_test)
   Debug.Assert FasCmdline.Commando = &H68 'That's Cond opcode
   
   On Error GoTo HandleCond_err
  
'    $213 32   0     3 Ld_INT8   push  00
'    $215 68   65    3           Cond If (0) Goto $25B Else .pop
   Dim cond_item_test
   cond_item_test = FasStack.Current
   
   
  '    Dim isCond_Last As Boolean
  '    isCond_Last = isCond_Last
      
      Dim isCond_First As Boolean
      isCond_First = cond_items.count = 0
      
      Static Cond_StartStack
      If isCond_First Then Cond_StartStack = FasStack.esp
      
      Dim hasRetValue As Boolean
      Dim bisCond_Last As Boolean
      bisCond_Last = isCond_Last(hasRetValue)
      
   ' find end of cond item
      With FileStream
        Dim FileStream_oldPos
        FileStream_oldPos = .Position
        ' .StorePos
         
          Dim cond_item_start&
          cond_item_start = FasCmdline.Parameters(1)
          
          .Move cond_item_start
          cond_item_start = .Position
      '    $256 57   9     4 GOTO      goto $264
      '  ;;; Cond item#0
      '    $25B A          3 Pop       pop dummy (decrease stack)
      '
      '    $25C 9    8     4 pu_Item   push "\ncond0"              (PRINC "\ncond0")
          
'        ' Pop_ Check PopA ' for empty body this is not there
'          Dim PopA
'          PopA = .int8
'          Debug.Assert PopA = &HA
'          .Move -1
          
         ' Move to GotoEnd
'          .Move -Goto_0x57_Size
          
        ' Goto_ CheckOpcode
'          Dim Goto_57
'          Goto_57 = .int8
'          Debug.Assert Goto_57 = &H57
          
          Dim cond_item_result_isEmpty As Boolean
          cond_item_result_isEmpty = Brancher.Goto_0x57_isNullJump(FromEnd:=True) 'Brancher.got(cond_item_end = 0)
          
          If Not cond_item_result_isEmpty Then
   
             
            ' Goto_ GetAdress
              Dim cond_item_end
              cond_item_end = Brancher.Goto_0x57_forward(DoGoto:=False, FromEnd:=True)  '.int32Sig
              
              
              'cond_item_end = cond_item_end + .Position
              'RangeCheck cond_item_end, .Length, 0, "Can't find the end of cond item", "HandleCond"
                 
           
         '    $25B A          3 Pop       pop dummy (decrease stack)
           ' Skip PopA
           ' Brancher.Skip_Pop_0a__NoPop
   
            ' To avoid decompile artefacts due to something on the stack
              FasStack.Current = ""
             
         '  ;;; Cond item#0
         '    $25B A          3 Pop       pop dummy (decrease stack)
         '    $25C 9    8     4 pu_Item   push "\ncond0"              (PRINC "\ncond0")
         '    $25F 35   1 15  4 ld_USUBR  PRINC 1 Params are above...
             
              ' Good place for a debugging Breakpoint
              ' on hasRetValue expected Stack delta is 0 (else -1) - same a value as bool->int
                Dim cond_item_result
                cond_item_result = Brancher.DoProgn(cond_item_end, Brancher.level, Not hasRetValue)
                
                
               ' A stack corruption occured during cond
               ' possible reasons:
               ' * this was no cond statement
               ' * hasRetValue heuristic failed
               If Cond_StartStack = FasStack.esp + IIf(hasRetValue, 0, 1) Then
'                 Stop
                 hasRetValue = Not hasRetValue
               End If
               
'               Debug.Assert Not ((FasStack.esp = 0) And (hasRetValue = True))
               If FasStack.esp = 0 Then hasRetValue = False
                 
                Dim dummy
                If hasRetValue Then dummy = FasStack.pop
 

         Else
            cond_item_result = ""
            cond_item_end = .Position
            FasStack.popIntoVoid
         End If
     ' .RestorePos
      .Position = FileStream_oldPos
      End With
   

   
   
   
   
    ' Make cond item
      Dim cond_Comment
      cond_Comment = "cond-item #" & cond_items.count
      FasCmdline.Disassembled = cond_Comment & FasCmdline.Disassembled
      cond_Comment = ""
      
      If cond_item_result_isEmpty Then _
         cond_item_result = TokenInlineComment("empty body")
      
      Dim cond_item
      cond_item = Join(Array( _
                     TokenComment(cond_Comment, 1), _
                     TokenOpen("", 1) & _
                     GetIndent(0) & cond_item_test & _
                     GetIndent(2) & cond_item_result, _
                     TokenClose("", 1) _
                     ), PartsSeperator)
      
      cond_items.add cond_item
   
   
   
  ' ... on First cond item
   If isCond_First Then
     ' Set cond_items = New Collection
      
    ' Store End of Cond
      cond_body_end = cond_item_end
      
     ' Debug.Assert Brancher.level = 0
      Brancher.level = 0
      
   End If
   
   If bisCond_Last Then
   
      Dim Cond_Full
      Cond_Full = Join(Array( _
                     TokenOpen("cond"), _
                     Join(CollectionToArray(cond_items), PartsSeperator), _
                     TokenClose("") _
                     ), PartsSeperator)
                     
     Set Cond_Full = make_ITEM(Cond_Full)
     
    ' clear items
      Set cond_items = New Collection
                     
      If hasRetValue Then
         FasStack.Current = Cond_Full
      Else
         Brancher.FasFile.outp Cond_Full
      End If
      
    ' Continue at the end of cond
      Debug.Assert FileStream.Position < cond_body_end ' Should be somewhere after current pos
    
      FileStream.Position = cond_body_end
      
   End If 'not isCond_Last
      
   
Exit Sub
HandleCond_err:
Stop
Resume


End Sub

Public Function isCond_Last(hasRetValue As Boolean)

   With FileStream
'Cond that is using a Retvalue
'   $9E 3         12         3 VALUE     Push value of [T]
'   $A1 68        6          3           Cond If (T) Goto $AC Else .pop

'   $A6 1                    4 ld NIL    Push nil
'   $A7 57        4          4 GOTO      goto $B0
'   $AC A                    3 Pop       pop dummy (decrease stack)                                   NIL
   
   
'   ;;; Cond RetValue (Terminal symbol)
'    $234 1          4 ld NIL    Push nil                   NIL
'    $235 57   0     4 GOTO      goto $23A
      Dim tmp
      tmp = .int16
      hasRetValue = (tmp = &H5701)
      isCond_Last = hasRetValue Or ((tmp And &HFF) = &H57)     '57_Goto <next>
      .Move -2
      
    ' Do has return value heuristic only if it's not the last cond item
      If Not isCond_Last Then
          
          .StorePos

          .Move -4
          Brancher.GotoTarget_Forward '           = branchTarget  (.Move .int32Sig)


          
            
            'Not There is Empty cond_itemBody
            'Debug.Assert .int8 = &HA:               .Move -1
            
      'Cond sample without RetValue
         '   $6E 0A             Pop       pop dummy (decrease stack)
         '       ^- Targeting this
         '   $6F 57  00000010   GOTO      goto $7E
         '   $74 0A             Pop       pop dummy (decrease stack)
            
      'Cond sample with RetValue
         '   $B6 09      0004    pu_Item   push "\n bAutoCAD 2018"
         '   $B9 57  00000004    GOTO      goto $C2
         '   $BE A        Pop    pop dummy (decrease stack)
            
             .Move -Goto_0x57_Size
             'Goto_0x57
            Debug.Assert .int8 = &H57:             .Move -1
            
            .Move -1
            hasRetValue = .int8 <> &HA:              .Move -1
             
          .RestorePos
         
      End If
   End With
End Function



'(and [expr ...])
Public Sub HandleAnd(FasCmdline As FasCommando) ', and_item_test)
   Debug.Assert FasCmdline.Commando = &H6A 'That's And opcode
   
   On Error GoTo HandleAND_err
  
'    $213 32   0     3 Ld_INT8   push  00
'    $215 6A   65    3           And If (0) Goto $25B Else .pop
   Dim and_item
   and_item = FasStack.pop
   
   Exit Sub
   
  '    Dim isCond_Last As Boolean
  '    isCond_Last = isCond_Last
      
      Dim isAnd_First As Boolean
      isAnd_First = and_item.count = 0
      
       Dim And_item_end&
       And_item_end = FasCmdline.Parameters(1)
         
     ' Good place for a debugging Breakpoint
     ' on hasRetValue expected Stack delta is 0 (else -1) - same a value as bool->int
       Dim and_item_result
       and_item_result = Brancher.DoProgn(And_item_end, Brancher.level)
       
      ' A stack corruption occured during cond
      ' possible reasons:
      ' * this was no cond statement
      ' * hasRetValue heuristic failed
'                 Debug.Assert Cond_StartStack = FasStack.esp + IIf(hasRetValue, 0, 1)
        
       'Dim dummy
       'If hasRetValue Then dummy = FasStack.pop

   
   
   
    ' Make and item
      Dim cond_item
      cond_item = Join(Array( _
                     TokenOpen("", 1) & _
                     GetIndent(0) & and_item & _
                     TokenClose("", 1) _
                     ), PartsSeperator)
      
      cond_items.add cond_item
   
   
   
  ' ... on First cond item
   If isAnd_First Then
     ' Set cond_items = New Collection
      
    '' Store End of Cond
     ' cond_body_end = cond_item_end
      
      Debug.Assert Brancher.level = 0
      Brancher.level = 0
      
   End If
   
'   If bisAnd_Last Then
   
      Dim and_Full
      and_Full = Join(Array( _
                     TokenOpen("and"), _
                     Join(CollectionToArray(And_Items), PartsSeperator), _
                     TokenClose("") _
                     ), PartsSeperator)
                     
     Set and_Full = make_ITEM(and_Full)
     
    ' clear items
      Set And_Items = New Collection
                     
      Brancher.FasFile.outp and_Full
      
      
'   End If
      
   
Exit Sub
HandleAND_err:
Stop
Resume


End Sub

