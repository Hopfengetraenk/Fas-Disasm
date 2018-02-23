VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VLXSpliter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const VLX_FILE_SIGNATURE$ = "VRTLIB-1"
        'Code from LISPLET.FSL
        '(  (cons :txt 1335) (cons :fas 1330) _
            (cons :DCL 1350) (cons :dvb 1340))
Enum VLX_ResTypeIds
   RT_LSP = &H0
   RT_PRV = &H4D8
   
   RT_TXT = &H537
   RT_FAS = &H532
   RT_DCL = &H546
   RT_DVB = &H53C
End Enum


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Log(ParamArray LogText())
   frmlog.listLog.AddItem Join(LogText, vbTab)
End Sub
Private Sub ReportExtractedFasFile(Filename$)
  
'   Dim newIndex&
'   newIndex = UBound(Filelist) + 1
'   ReDim Preserve Filelist(newIndex)
 '  Filelist(newIndex) = Filename
   Filelist.add Filename
   
End Sub


Public Function VLX_Split(strInputFilename As String) As Boolean
   On Error GoTo VLX_Split_err
   Dim InputFile As New FileStream
   Dim InputFilename As New ClsFilename
       InputFilename = strInputFilename
   
   Dim OutputFile As New FileStream
   Dim OutputFileName As New ClsFilename

   With InputFile
      .Create InputFilename.Filename
      
     'Log "vrtlib-id-string:", .FixedString(8)
      If .FixedString(Len(VLX_FILE_SIGNATURE$)) <> VLX_FILE_SIGNATURE$ Then _
         Err.Raise ERR_NO_VLX_FILE, , _
         "File doen't start with " & VLX_FILE_SIGNATURE & "!  This is propably no vlx-file."
      
      
     'Log "StreamSize:", H32(.int32)
      If .Length < .int32 Then _
         Err.Raise ERR_NO_VLX_FILE, , _
         "Datasize in VLX is greater than it's filesize"
      
      Dim Blocksize&, baseposition&
      
     'Set & Create outputdir
      OutputFileName.Path = InputFilename.Path & InputFilename.Name
      On Error Resume Next
         MkDir OutputFileName.Path
      On Error GoTo 0
      
      
    ' Process & Count all files in vlx
      Dim FileCount&
      For FileCount = 0 To &H7FFFFFFF
         baseposition = .Position
         
        'Get size of actual file
         Blocksize = .int32
         
    ' LOOP-EXIT Condition!
      If Blocksize = 0 Then Exit For
           
         'Log "blocksize:", H32(Blocksize)
         
        '=== Create File Name ====
        'Get type of file
         Dim resType As Integer: resType = .int16
         

         
        'Associate resType with File Extension
         Select Case resType
            Case RT_TXT: OutputFileName.Ext = "txt" '1335
            Case RT_FAS: OutputFileName.Ext = "fas" '1330
            Case RT_DCL: OutputFileName.Ext = "DCL" 'Dialog Control Language
            Case RT_DVB: OutputFileName.Ext = "dvb" '1340
   
            Case RT_PRV: OutputFileName.Ext = "prv" '1240
            Case RT_LSP: OutputFileName.Ext = "lsp"
            Case Else: OutputFileName.Ext = Hex(resType)
         End Select
         Log "resType:", Hex(resType), OutputFileName.Ext
         
        'GetFilename
         OutputFileName.Name = .FixedString(.int8)
        '
   
         
        'Add FileCount-number to filename if file already exists
         Dim FilesList As New Collection
         Dim FilesList_search_key$
         FilesList_search_key = OutputFileName.Name & OutputFileName.Ext
         On Error Resume Next
            FilesList (FilesList_search_key)
           'If Key was found add FileCount to name
            If Err.Number = 0 Then
               OutputFileName.Name = OutputFileName.Name & "_" & FileCount
            Else
               FilesList.add "", FilesList_search_key
            End If
            
         On Error GoTo 0
        
        'Notify Decompiler to process this file later
         If resType = RT_FAS Then ReportExtractedFasFile ( _
         InputFilename.Name & "\" & OutputFileName.Name & OutputFileName.Ext)
        
        '=== Get & Write OutData ====

        ' Create new File and write data

             OutputFile.Create OutputFileName.Filename, True
             OutputFile.FixedString(-1) = .FixedString((Blocksize - (.Position - baseposition)))
             OutputFile.CloseFile
   
        ' Align to fulldword (skip Fillbytes)
         .Position = (.Position + 3) And Not 3
           
       Next
       
      ' Log "vrtlib-id-string: ", .FixedString(8)
      ' Log "End_data: ", H32(.int32)
      .CloseFile
      
      frmlog.listLog.AddItem FileCount & " Files extracted to: " & OutputFileName.Path ', _
             vbInformation + vbOKOnly, "VLX-FileExtracter"
   '   ShellExecute " " & CurDir$ & "\" & OutputFilename.Name
   
   End With
  
  'set return value to show extaction successed
   VLX_Split = True

   Err.Clear
VLX_Split_err:
Select Case Err
   Case 0
   Case Else
      frmlog.listLog.AddItem Err.Description
End Select

End Function



