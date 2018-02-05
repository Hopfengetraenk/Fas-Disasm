Attribute VB_Name = "LspFile"
Option Explicit

      Const Byte_1A_EOF As Byte = &H1A
      Const Byte_0D_CR As Byte = &HD
      Const Byte_0A_LF As Byte = &HA

Private Const PROTECTED_LISP_FILE_SIGNATURE$ = "AutoCAD PROTECTED LISP file" ' & vbCrLf ' & Chr(&H1E)


Public Function LspFile_Decrypt(Filename$) As Boolean

   Dim LISP_FILE As New FileStream
   With LISP_FILE
      

       
       Dim LSP_InFileName As New ClsFilename
       LSP_InFileName = Filename
       
      .Create LSP_InFileName.Filename
      
       
       If isPROTECTED_LISP_FILE(LISP_FILE) Then
       
   
         Dim OutFile As New clsStrCat
         
         Dim key As Byte
         key = .int8
   
         Do
            Dim InData As Byte
            InData = .int8
            
         'If InData = Byte_1A_EOF Then Exit Do
            
               If (InData = Byte_1A_EOF) Or _
                  (InData = Byte_0D_CR) Then
                 
                 ' Skip all CR and EOF control Bytes
                 
               Else
               
               Dim Data_Out As Byte
               Data_Out = InData Xor key
               
               If (Data_Out = Byte_1A_EOF) Or _
                  (Data_Out = Byte_0D_CR) Then
                  Data_Out = InData
               End If
               
             ' Convert LF to LFCR
               If (Data_Out = Byte_0A_LF) Then
                  OutFile.ConcatByte Byte_0D_CR
               End If
               
               OutFile.ConcatByte Data_Out
               
               If Data_Out > &H80 Then
                  'Possible Error
'                  Stop
               End If
               
               Dim tmp
               tmp = InData
               tmp = tmp + tmp '= tmp << 1
               If tmp > 255 Then tmp = tmp - 255 '(tmp and 255)+1
               key = tmp
               
               
            End If
            
         Loop Until .EOS
         
   
         Dim LSP_OutFileName As New ClsFilename
         LSP_OutFileName = LSP_InFileName.Filename
         
         LSP_OutFileName.Name = LSP_OutFileName.Name & "_Dec"
        
         FileSave LSP_OutFileName.Filename, OutFile.value
         
         FrmMain.AddtoLog "Lisp File save to: " & LSP_OutFileName.Filename
       
      
         LspFile_Decrypt = True
       
       End If
       .Position = 0
       
   
      .CloseFile
   
   End With
End Function

Public Function isPROTECTED_LISP_FILE(LISP_FILE As FileStream)
   With LISP_FILE
    
      If .FixedString(Len(PROTECTED_LISP_FILE_SIGNATURE)) = PROTECTED_LISP_FILE_SIGNATURE Then
      
        isPROTECTED_LISP_FILE = False
        
        For i = 1 To 3
           If (.int8 = Byte_1A_EOF) Then
              isPROTECTED_LISP_FILE = True
              Exit For
           End If
        Next
        
      End If
      
    End With
    
End Function
