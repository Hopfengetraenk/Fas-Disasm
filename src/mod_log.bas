Attribute VB_Name = "mod_Log"
Option Explicit
Public Filename$
Public FileLog_Name

' length of txt log columns
Const TXTOUT_OPCODE_COL = 24
Const TXTOUT_DISASM_COL = 9
Const TXTOUT_DESCRIPT_COL = 63

'Offset_encode to encode Offset
'... to use as string
Public Function OffToVal(offset) As Long
   Dim tmp
   tmp = Trim(offset)
   
   'Debug.Assert tmp Like "$*"

   
   OffToVal = "&H" & Mid(tmp, 2)
End Function

'Offset_decode to decode Offset at a central place
'... to use as value
Public Function OffToStr(offset) As String
   OffToStr = "$" & Hex(offset)
End Function


Public Sub DoLog_OutputLine(outp As Log_OutputLine, LineBreaksCount)
   With outp
      Dim OutputLine As New clsStrCat
      
      Const TxtLog_ItemSeperator$ = " "
      Const TxtLog_ItemSeperator_len& = 1
      
    ' #1 Offset
    ' comment out when you like to compare output files later
      OutputLine.Concat BlockAlign_r(.offset, 6 + 1)
 '     Debug.Assert OffToVal(.offset) < &HFFFFFF
      OutputLine.Concat TxtLog_ItemSeperator
      
    ' #2 n #3 Command & Params
      OutputLine.Concat BlockAlign_r(.Command_Byte, 3)
      OutputLine.Concat TxtLog_ItemSeperator

      OutputLine.Concat .Params_Bytes ', 15)

      
    ' if it's to long OutputLine is not long - BlockAlign it
      Dim NotLongerThanThis&
      NotLongerThanThis = TXTOUT_OPCODE_COL + 1 ' +1 for the Spacer
      If (OutputLine.Length <= NotLongerThanThis) Then
         OutputLine.value = BlockAlign_l(OutputLine.value, NotLongerThanThis)
      Else
        ' else just let it run out of the column
        ' add a new line and an empty column
         
         
        ' OutputLine.Concat vbCrLf
         FileLog_Add OutputLine.value
         OutputLine.Clear
         OutputLine.Concat BlockAlign_l("", NotLongerThanThis)
         
      End If
      
      OutputLine.Concat TxtLog_ItemSeperator
      
   ' #4 Stack
      OutputLine.Concat BlockAlign_r(.Stack, 5)
      OutputLine.Concat TxtLog_ItemSeperator
      
      
    ' #5 n #6  ASM / Description
    
      OutputLine.Concat BlockAlign_l(.DisASM, TXTOUT_DISASM_COL)
      Inc NotLongerThanThis, TXTOUT_DISASM_COL + TxtLog_ItemSeperator_len
      OutputLine.Concat TxtLog_ItemSeperator

      OutputLine.Concat BlockAlign_l(.Description, TXTOUT_DESCRIPT_COL)
      Inc NotLongerThanThis, TXTOUT_DESCRIPT_COL + TxtLog_ItemSeperator_len
      OutputLine.Concat TxtLog_ItemSeperator
      
    ' Crop OutputLine if it's to long (so it fit's in the column
    ' NotLongerThanThis = TXTOUT_OPCODE_COL + TXTOUT_DISASM_COL
    ' If (OutputLine.Length <= NotLongerThanThis) Then

      OutputLine.value = BlockAlign_l(OutputLine.value, NotLongerThanThis)
      
      
   ' #7 Decompiled
      OutputLine.Concat .DeCompiled
      
      
    ' Add linebreaks
 '     Dim LineBreaksCount
 '     Output_GetLineBreaks .Description, LineBreaksCount
      
      For i = 1 To LineBreaksCount
         OutputLine.Concat vbCrLf
      Next

      FileLog_Add OutputLine.value


   End With
End Sub


Public Sub Output_GetLineBreaks(DisASM, LineBreaksCount)

    ' Get line breaks in Disasm
      Dim lenBefore&
      lenBefore = Len(DisASM)
      DisASM = Replace(DisASM, vbCrLf, "")
      
      LineBreaksCount = (lenBefore - Len(DisASM))
      LineBreaksCount = LineBreaksCount \ Len(vbCrLf)
      
      
'    ' Newline if ESP=0
'      Static lastStackitem
'      If File.FasStack.ESP < lastStackitem Then
'         LineBreaksCount = LineBreaksCount + 1
'      'Else
'      End If
'      lastStackitem = File.FasStack.ESP

End Sub

Public Sub FileLog_open()
   On Error Resume Next

   FileLog_close

   FileLog_Name = Filename & ".txt"
   Open FileLog_Name For Output As 1
End Sub

Public Sub FileLog_Add(TextLine$)
   On Error Resume Next

   Print #1, TextLine
End Sub
Public Sub FileLog_close()
   On Error Resume Next

'if you stop here you'd
'proably enabled stop
'on all Errors in the VB-IDE
   Dim isEmptyFile As Boolean
   isEmptyFile = LOF(1) = 0
   Close #1
   
   If isEmptyFile Then _
      Kill FileLog_Name
   
End Sub


Public Sub SaveDecompiled()

   Dim lsp_Filename
   lsp_Filename = Filename & "_.lsp"
   
   Const ERR_FileNotFound = 53
   On Error Resume Next
   Kill lsp_Filename
   If Err And (Err <> ERR_FileNotFound) Then _
      FrmMain.AddtoLog _
         "Can't delete " & lsp_Filename & _
         " ERR:" & Err.Description

   
   Open lsp_Filename For Output Shared As 2
   Dim item, i
 ' note: using 'for each' here might dump garbage since Storage might be bigger than .esp
   For i = 0 To FrmMain.LispFileData.esp
      Print #2, FrmMain.LispFileData.Storage(i)
   Next
   Close #2

End Sub

