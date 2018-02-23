VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public Filename As String

Public mvarhFile As Long
Private mvarByteToBeRead As Long 'lokale Kopie
Private mvarMaxPosition As Long 'lokale Kopie
Public bIsTemporaryFile As Boolean
Public Readonly As Boolean

Public DisableMoveOnRead As Boolean
Public DisableMoveOnWrite As Boolean


Private CreationDisposition&
Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_SHARE_DELETE = &H4

Private Const CREATE_ALWAYS = 2
Private Const CREATE_NEW = 1
Private Const OPEN_ALWAYS = 4
Private Const OPEN_EXISTING = 3
Private Const TRUNCATE_EXISTING = 5

Private FileFlagsAndAttributes&
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Private Const FILE_FLAG_NO_BUFFERING = &H20000000
Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Declare Function ReadFilePtr Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFileLong Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByRef lpBuffer As Long, ByVal nNumberOfBytesToRead As Integer, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFileDouble Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByRef lpBuffer As Double, ByVal nNumberOfBytesToRead As Integer, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function ReadFileInt64 Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByRef lpBuffer As Currency, ByVal nNumberOfBytesToRead As Integer, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (src As Any, src As Any, ByVal Length&)
Private Declare Sub MemCopyStrToLng Lib "kernel32" Alias "RtlMoveMemory" (src As Long, ByVal src As String, ByVal Length&)
Private Declare Sub MemCopyLngToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal src As String, src As Long, ByVal Length&)
Private Declare Sub MemCopyLngToInt Lib "kernel32" Alias "RtlMoveMemory" (src As Long, ByVal src As Integer, ByVal Length&)
'Private Declare Sub MemCopyStrToInt Lib "kernel32" Alias "RtlMoveMemory" (src As Any, Src As String, ByVal length&)
'Private Declare Sub MemCopyStrToByte Lib "kernel32" Alias "RtlMoveMemory" (src As Any, Src As String, ByVal length&)

Private Enum SeekType
   FILE_BEGIN = 0
   FILE_CURRENT = 1
   FILE_END = 2
End Enum

Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

' An attempt was made to move the file pointer before the beginning of the file.
Private Const ERROR_NEGATIVE_SEEK& = 131

' The file pointer cannot be set on the specified device or file.
Private Const ERROR_SEEK_ON_DEVICE& = 132


Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long

Private retval As Long
Private bytesRead As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long


Private mData As New StringReader
Private mMemOnlyMode As Boolean

Private mPosOld&


''FileTime
'Private Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCreationTime As FileTime, ByRef lpLastAccessTime As FileTime, ByRef lpLastWriteTime As FileTime) As Long
'Private Type FileTime
'   dwLowDateTime As Long
'   dwHighDateTime As Long
'End Type
'
'
'
'Public Property Get LastAccessDate() As Long
'   Dim CreationTime As FileTime
'   Dim LastAccessTime As FileTime
'   Dim LastWriteTime As FileTime
'
'   GetFileTime hFile, CreationTime, LastAccessTime, LastWriteTime
'
'   LastAccessDate = LastAccessTime.dwHighDateTime
'
'End Property

Public mStorePos As Long
Private StorePosLock As Long

Public Sub RestorePos()
   If StorePosLock > 0 Then _
      Dec StorePosLock _
   Else _
      Err.Raise vbObjectError, "FileStream::RestorePos", _
               "There is no location stored. Use StorePos() to store one." & _
               "Note: You can call RestorePos() only once."
      
   Position = mStorePos
End Sub

Public Sub StorePos()
   If StorePosLock = 0 Then _
      Inc StorePosLock _
   Else _
       Err.Raise vbObjectError, "FileStream::StorePos" _
               , "There is already one location stored - call RestorePos() first so StorePos() will get free to use!"

      
   mStorePos = Position
End Sub



Private Function StrtoLng(ByVal value$) As Long
   MemCopyStrToLng StrtoLng, value, 4
End Function

Private Function LngtoStr(ByRef value&) As String
   Dim tmp$
   tmp = Space(4)
   MemCopyLngToStr tmp, value, 4
   LngtoStr = tmp
End Function


'Public Property Let EOS(ByVal vData As Boolean)
'    mvarEOS = vData
'End Property

Public Property Get EOS() As Boolean
    EOS = Me.Length = Me.Position
'    Debug.Assert mData.EOS = EOS
End Property

'// Set the EOS at the current position (= .position)
Public Sub setEOS()

   If mMemOnlyMode Then
      mData.Truncate
   Else
   
      SetEndOfFile hFile
   End If

   
   mvarMaxPosition = Position

End Sub

'Public Property Let FixedStringW(ByVal Length As Long, ByRef vData As String)
'   If Length = -1 Then Length = Len(vData)
'   If WriteFile(hFile, vData, Length, bytesRead, 0) <= 0 Then Err.Raise vbObjectError, , "Error in Property Let fixedString: Can't write data. Filestream::Readonly=" & Readonly
'End Property

Public Property Get FixedStringW(ByVal Length As Long) As String

   If mMemOnlyMode Then Stop ' Not implemented yet
   
   
   If Length = -1 Then
      Length = (Me.Length - Me.Position)
   End If
 
 ' create Buffer
   FixedStringW = Space(Length)
   
 ' Read Fixed String
   retval = ReadFilePtr(hFile, StrPtr(FixedStringW), Length * 2, bytesRead, 0)
   If retval < 0 Then Err.Raise vbObjectError, , "ReadFileString failed."
   
   If DisableMoveOnRead Then SeekToPosition bytesRead, FILE_CURRENT

End Property






Public Property Let FixedString(ByVal Length As Long, ByRef vData As String)
   
   If mMemOnlyMode Then
      mData.FixedString(Length) = vData
   Else
   
   
   If Length = -1 Then Length = Len(vData)
 ' deal with the case of negative FileOffsets
   If Me.Position < 0 Then
    ' shorten string
    MsgBox "TODO: Handle write data at an negative fileoffset."
    
'      Length = Length + Position
      
'      If Length >= 0 Then
       ' set to beginning of file - so readfile has a chance to succeed
 '        Position = 0
 '     Else
       ' whoops after FileRead we'll still 'before the beginning'
       ' so just move 'virtually' forward and return an empty string
'         Position = -Length
'         FixedString = ""
'         Exit Property
         
'      End If

   End If
   
   
   If WriteFile(hFile, vData, Length, bytesRead, 0) <= 0 Then Err.Raise vbObjectError, , "Error in Property Let fixedString: Can't write data. Filestream::Readonly=" & Readonly
   
   
   If DisableMoveOnWrite Then SeekToPosition bytesRead, FILE_CURRENT
  
   
   End If


End Property

Public Property Get FixedString(ByVal Length As Long) As String
   If mMemOnlyMode Then
      FixedString = mData.FixedString(Length)
   Else


   ' 'Length = -1' means read till end
   If Length = -1 Then
   ' deal with the case of negative FileOffsets
     If Me.Position < 0 Then
      ' seek to start
        Me.Position = 0
      ' 'select' whole file
        Length = Me.Length
      Else
      ' 'select' from current pos to end of file
         Length = (Me.Length - Me.Position)
      End If
   End If
 
 ' deal with the case of negative FileOffsets
   If Me.Position < 0 Then
    ' shorten length
      Length = Length + Position
      
      If Length >= 0 Then
       ' set to beginning of file - so readfile has a chance to succeed
         Position = 0
      Else
       ' whoops after FileRead we'll still 'before the beginning'
       ' so just move 'virtually' forward and return an empty string
         Position = -Length
         FixedString = ""
         Exit Property
         
      End If
   End If
   
   
 ' create Buffer
   On Error Resume Next
   FixedString = Space(Length)
   
   
 ' Read Fixed String
   retval = ReadFile(hFile, FixedString, Length, bytesRead, 0)
   If retval < 0 Then Err.Raise vbObjectError, , "ReadFileString failed."
   
   
   
 ' Limit buffer incase the file is smaller
   FixedString = Left(FixedString, bytesRead)
   
   If Err Then
      Dim tmp
      tmp = "@FileStream::FixedString   Error allocating " & Length \ 1024 & " KB."
      MsgBox tmp, vbCritical, Err.Description
      
      On Error GoTo 0
      Err.Raise vbObjectError, , tmp
      
      
   End If
   
   On Error GoTo 0
   
   If DisableMoveOnRead Then SeekToPosition bytesRead, FILE_CURRENT
   
   End If
End Property

Public Property Let char(ByRef vData As String)

   If mMemOnlyMode Then
      mData.FixedString(1) = vData
      
   Else
   
      FixedString(1) = vData
      
   End If
End Property

Public Property Get char() As String

   If mMemOnlyMode Then
      char = mData.FixedString(1)
   Else
   
      char = FixedString(1)
   End If
End Property


Public Sub FindBytes(ParamArray Bytes())

   If mMemOnlyMode Then Stop ' Not implemented yet
   
   
   Dim i
   i = LBound(Bytes)
   
   With Me
      Do
         If .int8 = Bytes(i) Then
            If i >= UBound(Bytes) Then Exit Do
            i = i + 1
         Else
            i = LBound(Bytes)
         End If
      Loop Until .EOS
   End With
End Sub

Public Function FindString&(StringToFind$, Optional SearchBackwards As Boolean = False, Optional VbCompareMethod As VbCompareMethod = vbBinaryCompare)

   If mMemOnlyMode Then Stop ' Not implemented yet
   
   
   Dim tmp$, oldPos&
   oldPos = Position
      tmp = FixedString(-1)
      
      If SearchBackwards Then
         FindString = InStrRev(tmp, StringToFind, , VbCompareMethod) - 1
      Else
         FindString = InStr(1, tmp, StringToFind, VbCompareMethod) - 1
      End If
      
      If FindString = -1 Then
         Position = oldPos
'         FindString = 0
      Else
         Position = FindString + oldPos
      End If
'   Dim i
'   i = 1
'
'   With Me
'      Do
'         If .FixedString(1) = Mid(StringToFind, i, 1) Then
'            If i >= Len(StringToFind) Then Exit Do
'            i = i + 1
'         Else
'            i = 1
'         End If
'      Loop Until .EOS
'   End With
'   Position = (Position - Len(StringToFind)) And Not (EOS)
'   FindString = Position
End Function

Public Function FindStrings(StringToFind$, Optional StopIfMoreThan& = &H7FFFFFFF) As Collection

   If mMemOnlyMode Then Stop ' Not implemented yet
   
   Set FindStrings = New Collection
   
   Dim tmp$, oldPos&
   oldPos = Position
   Position = 0
   tmp = FixedString(-1)
   
   Dim offset&, item
   offset = 0
   For Each item In Split(tmp, StringToFind, StopIfMoreThan)
      
      Inc offset, Len(item)
      FindStrings.add offset
      
      Inc offset, Len(StringToFind)
   Next
   FindStrings.Remove FindStrings.count
   
 ' Restore old position if not found
   If FindStrings.count = 0 Then Position = oldPos
   
End Function



Public Function getTerminatedString(ParamArray TerminatorStrings()) As String


   If mMemOnlyMode Then
      Debug.Assert TerminatorStrings(0) = Chr(0)
      getTerminatedString = mData.zeroString()
   
   
   
   Else
   
  'For i = LBound(TerminatorStrings) To UBound(TerminatorStrings)
  'If Len(TerminatorString) = 0 Then Exit Function...

  Dim TerminatorStringsMatchIndexes
  ReDim TerminatorStringsMatchIndexes(UBound(TerminatorStrings))
  
  Dim value As String
  value = " "

  Dim i&
  
  '... because it's not implemented for that methode
  Debug.Assert DisableMoveOnRead = False

' Begin of FileRead-Loop
  Do
   ' Fill buffer
     If ReadFile(hFile, value, 1, bytesRead, 0) <= 0 Then Err.Raise vbObjectError, , "ReadFileString failed."

   ' if no byte was read we reached the End of File
     If bytesRead = 0 Then
      ' Clear getTerminatedString
        getTerminatedString = ""
        Exit Function
     End If
   
   ' append char to String
     getTerminatedString = getTerminatedString & value
   
     For i = LBound(TerminatorStrings) To UBound(TerminatorStrings)
    
      ' If char of the string does not match ...
        If value <> Mid$(TerminatorStrings(i), TerminatorStringsMatchIndexes(i) + 1, 1) Then
         
         '... reset stringIndexMatchPointer
          TerminatorStringsMatchIndexes(i) = 0
        
        Else
          
        '... increase stringIndexMatchPointer
          TerminatorStringsMatchIndexes(i) = TerminatorStringsMatchIndexes(i) + 1
        ' does String fully match ?
          If TerminatorStringsMatchIndexes(i) >= Len(TerminatorStrings(i)) Then
           
           ' Cut off matchstring
             getTerminatedString = Left(getTerminatedString, Len(getTerminatedString) - Len(TerminatorStrings(i)))
           
           ' exit FileRead-Loop
             Exit Do
             
          End If
        
        End If
     Next
 
 ' end of FileRead-Loop
   Loop Until Me.EOS
     

   End If

End Function

Public Property Let zeroString(ByRef vData As String)
   Stop
End Property


Public Property Get zeroString() As String
   
   zeroString = Me.getTerminatedString(Chr(0))


'    ' ... read Zero Terminated String
'      Dim EOS As Boolean, value As Byte
'      Loop
'
'      ' Fill buffer
'        retVal = ReadFileString(hFile, value, 1, bytesRead, 0)
'
'        If retVal > 0 Then
'
'         ' If we reached the end of string...
'           If value = 0 Then
'            ' ...yes exit loop
'              Exit Do
'           Else
'            '...no - append char to String
'              zeroString = zeroString & Chr(value)
'           End If
'
'      ' if no byte was read we reached the End of File
'        ElseIf retVal = 0 Then
'           EOS = True
'        Else
'           Err.Raise vbObjectError, , "ReadFileString failed."
'        End If
'
'      Do Until EOS
'   End If
'    --- Faster Version, but inproper stream implementiation (because of rewind)
'    Dim bytesRead&
'    Dim strlen&, GotString&, EOS As Boolean
'    ', value As Byte
'    Const readBuffSize& = 256
'    Dim readBuffer$
'
'
'  ' Create buffer
'    readBuffer = String(readBuffSize, Chr(0))
'
'    Loop
'
'    ' Fill buffer
'      bytesRead = ReadFileString(hFile, readBuffer, readBuffSize, bytesRead, 0)
'
'    ' if the buffer wasn't filled completely we reached the EOS(End Of File)
'      EOS = bytesRead <> readBuffSize
'
'    ' Cut off buffer if the buffer wasn't filled
'      If EOS Then readBuffer = Left(readBuffer, bytesRead)
'
'    ' Find End of String
'      GotString = InStr(0, readBuffer, Chr(0))
'
'    ' if GotString=true (<>0)...
'      If GotString Then
'       ' Append Readbuffer to GetString
'         GetString = GetString & Left(readBuffer, GotString)
'      Else
'       ' Append whole Readbuffer to GetString
'         GetString = GetString & readBuffer
'      End If
'
'    Do Until EOS Or GotString
'
'   'rewind to end of string
'    position = position - (readBuffSize - GotString)

End Property



Public Property Let int32(ByRef vData As Long)

   If mMemOnlyMode Then Stop ' Not implemented yet

    Dim retval&
    Dim bytesWritten&
    Dim tmp$
    tmp = LngtoStr(vData)
    retval = WriteFile(hFile, tmp, 4, bytesWritten, 0)
    If bytesWritten <> 4 Then
      Err.Raise vbObjectError, "", "Let_int32: WriteFile failed! Readonly=" & CBool(vbReadOnly)
    End If

   If DisableMoveOnWrite Then SeekToPosition bytesWritten, FILE_CURRENT

End Property


Public Property Get int32() As Long
   
   
   If mMemOnlyMode Then
   
      int32 = mData.int32
   Else

'    Dim bytesRead&, value As Long
'    bytesRead = ReadFileLong(hFile, value, 4, bytesRead, 0)
'    int32 = value

  Dim retval&, bytesRead&
  retval = ReadFileLong(hFile, int32, 4, bytesRead, 0)
  If bytesRead <> 4 Then MsgBox "[File::int32] Only '" & bytesRead & "' bytes read instead of 4! RetVal: " & retval, vbCritical
  
  End If
  
' values are maybe not unsigned plz Test
  Debug.Assert int32 >= 0
  
  If DisableMoveOnRead Then SeekToPosition bytesRead, FILE_CURRENT

  
End Property

Public Property Get int32Sig() As Long

   If mMemOnlyMode Then
   
      int32Sig = mData.int32Sig
   Else


  Dim retval&, bytesRead&
  retval = ReadFileLong(hFile, int32Sig, 4, bytesRead, 0)
  If bytesRead <> 4 Then MsgBox "[File::int32] Only '" & bytesRead & "' bytes read instead of 4! RetVal: " & retval, vbCritical
  
  End If
  
  ' values are maybe not signed plz Test
  Debug.Assert int32Sig < &H7FFFFFFF

  If DisableMoveOnRead Then SeekToPosition bytesRead, FILE_CURRENT


  
End Property

Public Property Let int16(ByRef vData As Long)
      Stop
End Property


Public Property Get int16Sig() As Integer
    Dim value As Long
    value = int16()
   
    ' 54298
    '  int16sig = value Or -(value And &H8000&)  '-32768 '&H8000
    If value And &H8000& Then
       int16Sig = (value And &H7FFF) Or &H8000&
    Else
       int16Sig = value
    End If
End Property

Public Property Get int16() As Long
    

    Dim bytesRead&, value As Long
    
    If mMemOnlyMode Then 'Stop ' Not implemented yet
       value = mData.int16
    Else
       bytesRead = ReadFileLong(hFile, value, 2, bytesRead, 0)
    
    End If
    
    If DisableMoveOnRead Then SeekToPosition bytesRead, FILE_CURRENT
   
    int16 = value
    
End Property

Public Property Let DoubleValue(ByRef vData As Double)

Stop
'    Dim bytesWritten&
'    Dim tmp$
'    tmp = LngtoStr(vData)
'
'    WriteFile hFile, tmp, 8, bytesWritten, 0



End Property


Public Property Get DoubleValue() As Double

  If mMemOnlyMode Then Stop ' Not implemented yet

  Dim bytesRead
  bytesRead = ReadFileDouble(hFile, DoubleValue, 8, bytesRead, 0)
  If DisableMoveOnRead Then SeekToPosition bytesRead, FILE_CURRENT

End Property


Public Property Get int64Value() As Currency

  If mMemOnlyMode Then Stop ' Not implemented yet
   
  Dim bytesRead
  bytesRead = ReadFileInt64(hFile, int64Value, 8, bytesRead, 0)
  If DisableMoveOnRead Then SeekToPosition bytesRead, FILE_CURRENT

End Property




'////////////////////////////////////////
'//  Let position Property
Public Property Let Position(ByVal vData As Long)

   If mMemOnlyMode Then
      mData.Position = vData
   Else
   
   SeekToPosition vData, FILE_BEGIN
   
   End If
End Property
'//  Get position Property
Public Property Get Position() As Long
Attribute Position.VB_UserMemId = 0

   If mMemOnlyMode Then
      Position = mData.Position
   Else

   Position = SetFilePointer(hFile, 0, 0, FILE_CURRENT)
   If Position < 0 Then
'      MsgBox "Error on GetFilePointer, FilePointer is before beginning: " & Position, vbCritical
'      Me.Position = 0 'SetFilePointer hFile, 0, 0, FILE_BEGIN
   End If
'  position = mvarposition

   End If
End Property


'////////////////////////////////////////
'//  Let PositionEOFRelated Property
Public Property Let PositionEOFRelated(ByVal vData As Long)
   If mMemOnlyMode Then
      Err.Raise 0, , "Not supported/implemented yet for MemMode!"
   Else


   SeekToPosition vData * -1, FILE_END
  
  'For some stupid reason it requires SetFilePointer( with FILE_BEGIN)
  ' or next fileRead will fail
   Position = Position
   End If
End Property


Private Sub SeekToPosition(offset&, SeekType As SeekType)
   If mMemOnlyMode Then
      mData.Position = offset
      Debug.Assert SeekType = FILE_BEGIN
   Else

   Dim retval&
   retval = SetFilePointer(hFile, offset, 0, SeekType)
   If retval <> ERROR_NEGATIVE_SEEK Then
   
      If Err.LastDllError Then
         MsgBox "Error on SetFilePointer, Offset: " & offset & "  SeekType: " & SeekType & vbCrLf & _
                "LastDllError: " & Err.LastDllError, _
                vbCritical
      End If
      
   Else
      mvarMaxPosition = Max(mvarMaxPosition, retval)
   End If
   
   End If
End Sub


''////////////////////////////////////////
''//  Let ByteToBeRead Property
'Public Property Let ByteToBeRead(ByVal vData As Long)
'    mvarByteToBeRead = vData
'End Property
''//  Get ByteToBeRead Property
'Public Property Get ByteToBeRead() As Long
'    If (mvarByteToBeRead < 0) Or (mvarByteToBeRead > Length) Then mvarByteToBeRead = Length
'    ByteToBeRead = mvarByteToBeRead
'End Property



'////////////////////////////////////////
'//  Set File Handle Property
Private Property Let hFile(ByVal vData As Long)
   
  'If invalid file handle is to be set...
   If vData = -1 Then
     'Close File
      CloseHandle mvarhFile
      
    ' DeleteFile if it is opened as Temporary
      If bIsTemporaryFile Then
         DeleteFile Filename
      End If
      
   End If
 
 ' Store Filehandle
   mvarhFile = vData

End Property
'//  Get File Handle Property

Public Property Get hFile() As Long
   
   Dim mShare
   mShare = FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE
   
 ' If no file handle(mvarhFile = 0) get one.
   If mvarhFile = -1 Then
     'Open file as with new file handle
      mvarhFile = CreateFile(Filename, _
                              GENERIC_READ Or (GENERIC_WRITE And Not (Readonly)), _
                              mShare, _
                              0, CreationDisposition, _
                              FileFlagsAndAttributes, 0)
      If mvarhFile = -1 Then
      
         On Error Resume Next
         Dim tmptxt$
         tmptxt = IIf(GetAttr(Filename) And vbReadOnly, "File is write protected.", "File is in use.")
         
         If Err Then tmptxt = Err.Description
         On Error GoTo 0
         Err.Raise ERR_OPENFILE, , "Open file for read" & _
                  IIf(Readonly, "", "/write") & " shareflags: 0x" & H32(mShare) & _
                  " access fail. " & tmptxt
         '" & Filename
      End If
   End If

 ' return Filehandle
   hFile = mvarhFile
  
 ' Set FileSize
   mvarMaxPosition = GetFileSize(mvarhFile, 0)
   
End Property

'////////////////////////////////////////
'//  Get Length Property
Public Property Get Length() As Long

   If mMemOnlyMode Then
      mvarMaxPosition = mData.Length
   Else
'->TODO: Cache lenght of file in a variable)
    mvarMaxPosition = GetFileSize(hFile, 0)
   End If
    
    Length = mvarMaxPosition
'    Length = Max(mvarMaxPosition, GetFileSize(mvarhFile, 0), Me.Position)
'    Length = Max(mvarMaxPosition, FileLen(FileName), Me.Position)

End Property





'////////////////////////////////////////
'//  Let Data Property
Public Property Let int8(vData As Long) 'Offset&, lenght&, ByRef vData As String)

   If mMemOnlyMode Then
      mData.int8 = vData
   Else
    
    
    
    Dim bytesWritten&
    Dim tmp$
    tmp = LngtoStr(vData)

    
    WriteFile hFile, tmp, 1, bytesWritten, 0
    Debug.Assert bytesWritten = 1
    
    If DisableMoveOnWrite Then SeekToPosition bytesWritten, FILE_CURRENT

    End If
End Property
'//  Get Data Property
Public Property Get int8() As Long '(Offset&, lenght&)
   If mMemOnlyMode Then
      int8 = mData.int8
   Else

    Dim bytesRead&
    Dim tmp$
    tmp = " "
    retval = ReadFile(hFile, tmp, 1, bytesRead, 0)
    If bytesRead <> 1 Then
      Debug.Print Err.LastDllError
      Stop
    End If
    int8 = Asc(tmp)
    
    If DisableMoveOnRead Then SeekToPosition -bytesRead, FILE_CURRENT


    End If

End Property
          
'////////////////////////////////////////
'//  Let Data Property
Public Property Let int8Sig(vData As Integer) 'Offset&, lenght&, ByRef vData As String)

   If mMemOnlyMode Then
   Stop
      'mData.int8Sig = vData
   Else
    
    Dim bytesWritten&
    Dim tmp$
    tmp = LngtoStr(CLng(vData))

    WriteFile hFile, tmp, 1, bytesWritten, 0
    If DisableMoveOnWrite Then SeekToPosition bytesWritten, FILE_CURRENT
    
    End If
End Property
'//  Get Data Property
Public Property Get int8Sig() As Integer '(Offset&, lenght&)
   If mMemOnlyMode Then
      int8Sig = mData.int8Sig
   Else

      Dim tmp
      tmp = int8()

'    int8Sig = int8Sig Or -(int8Sig And &H80)
   ' convert unsigned char to signed char
     If tmp > &H7F Then
        int8Sig = (tmp And &H7F) Or &HFF80
     End If

    End If

End Property
          



'////////////////////////////////////////
'//  Create
Public Sub Create(Filename As String, Optional bDeleteExistingFile As Boolean, Optional bTemporaryFile As Boolean, Optional Readonly As Boolean)
   CloseFile
   Me.Filename = Filename
   Me.Readonly = Readonly
   If bDeleteExistingFile Then
      CreationDisposition = CREATE_ALWAYS
   Else
      CreationDisposition = OPEN_EXISTING
   End If
   
   bIsTemporaryFile = bTemporaryFile
   FileFlagsAndAttributes = FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_SEQUENTIAL_SCAN
   
   If bIsTemporaryFile Then FileFlagsAndAttributes = FileFlagsAndAttributes Or FILE_ATTRIBUTE_TEMPORARY
   
End Sub

'////////////////////////////////////////
'//  CloseFile
Public Sub CloseFile()
   hFile = -1
End Sub

Private Sub Class_Initialize()
 ' Defaults
   mvarhFile = -1
End Sub

Private Sub Class_Terminate()
   CloseFile
End Sub

''////////////////////////////////////////
''//  Get Data Int
'Public Function GetInt() As Long
'    Dim bytesRead&
'    bytesRead = ReadFileInt(hFile, GetInt, 2, bytesRead, 0)
'End Function
''//  Get Data Long
'Public Function GetLong() As Long
'    Dim bytesRead&
'    bytesRead = ReadFileLong(hFile, GetLong, 4, bytesRead, 0)
'End Function
''//  Get Data String
'Public Function GetString(Optional lenght& = -1) As String
'   Dim retVal&
'
' ' if length was specified...
'   If lenght = -1 Then
'    ' create Buffer
'      GetString = Space(lenght)
'    ' ... read Fixed String
'      retVal = ReadFileString(hFile, GetString, lenght, bytesRead, 0)
'      If retVal < 0 Then Err.Raise vbObjectError, , "ReadFileString failed."
'
'   Else
'    ' ... read Zero Terminated String
'      Dim EOF As Boolean, value As Byte
'      Loop
'
'      ' Fill buffer
'        retVal = ReadFileString(hFile, value, 1, bytesRead, 0)
'
'        If retVal > 0 Then
'
'         ' If we reached the end of string...
'           If value = 0 Then
'            ' ...yes exit loop
'              Exit Do
'           Else
'            '...no - append char to String
'              GetString = GetString & Chr(value)
'           End If
'
'      ' if no byte was read we reached the End of File
'        ElseIf retVal = 0 Then
'           EOF = True
'        Else
'           Err.Raise vbObjectError, , "ReadFileString failed."
'        End If
'
'      Do Until EOF
'   End If
'
''    --- Faster Version, but inproper stream implementiation (because of rewind)
''    Dim bytesRead&
''    Dim strlen&, GotString&, EOF As Boolean
''    ', value As Byte
''    Const readBuffSize& = 256
''    Dim readBuffer$
''
''
''  ' Create buffer
''    readBuffer = String(readBuffSize, Chr(0))
''
''    Loop
''
''    ' Fill buffer
''      bytesRead = ReadFileString(hFile, readBuffer, readBuffSize, bytesRead, 0)
''
''    ' if the buffer wasn't filled completely we reached the EOF(End Of File)
''      EOF = bytesRead <> readBuffSize
''
''    ' Cut off buffer if the buffer wasn't filled
''      If EOF Then readBuffer = Left(readBuffer, bytesRead)
''
''    ' Find End of String
''      GotString = InStr(0, readBuffer, Chr(0))
''
''    ' if GotString=true (<>0)...
''      If GotString Then
''       ' Append Readbuffer to GetString
''         GetString = GetString & Left(readBuffer, GotString)
''      Else
''       ' Append whole Readbuffer to GetString
''         GetString = GetString & readBuffer
''      End If
''
''    Do Until EOF Or GotString
''
''   'rewind to end of string
''    position = position - (readBuffSize - GotString)
'
'End Function
''////////////////////////////////////////
''//  Set Data Int
'
'Private Function getLen(Var As Variant) As Long
'   Select Case TypeName(Var)
'   Case "Byte":
'      getLen = 1
'   Case "Integer":
'      getLen = 2
'   Case "String":
'      getLen = Len(Var)
'   Case "Long":
'      getLen = 4
'   Case Else
'      Err.Raise vbObjectError, , , "getLen - unknown VarType"
'   End Select
'End Function
'
''Private Sub SetFilePosition()
''   Dim retval&
''   retval = SetFilePointer(hFile, Offset, 0, FILE_BEGIN)
''   If retval = -1 Then MsgBox "Error on SetFilePointer", vbCritical
''End Sub


Public Sub Move(BytesToMoveFromCurPos&)
   
   Position = Position + BytesToMoveFromCurPos&
   mvarMaxPosition = Max(mvarMaxPosition, Position)
   
'   Dim RetVal&
'   RetVal = SetFilePointer(hFile, BytesToMoveFromCurPos, 0, FILE_CURRENT)
'   If RetVal = -1 Then
'      MsgBox "Error on moving FilePointer Position: " & Position & " +" & BytesToMoveFromCurPos, vbCritical
'   Else
'      mvarMaxPosition = Max(mvarMaxPosition, RetVal)
'   End If
End Sub


Public Property Let data(ByRef vData As String)
   Position = 0
   FixedString(-1) = vData
   Position = 0
End Property

Public Property Get data() As String
   Position = 0
   data = FixedString(-1)
   Position = 0
End Property



Public Property Get MemOnlyMode() As Variant
   MemOnlyMode = mMemOnlyMode
End Property

Public Property Let MemOnlyMode(ByVal vNewValue As Variant)
   
   If vNewValue Then
      MemOnlyOn
   Else
      MemOnlyOff
   End If

End Property



Private Sub MemOnlyOn()
'   Debug.Assert mMemOnlyMode = False
   
   If mMemOnlyMode = True Then
      Debug.Print "MemOnly Mode already enabled"
      Exit Sub
   End If
   mMemOnlyMode = False
   
 ' Store old Position
   mPosOld = Position
   
   MemReFill
   
   
' Restore old position
  Position = mPosOld
  
   
End Sub
Private Sub MemOnlyOff()
   Debug.Assert mMemOnlyMode = True
 
 ' Store old Position
   mPosOld = Position
   
   MemFlush
   
' Restore old position
  Position = mPosOld
End Sub

Private Sub MemReFill()
 'Read whole file into buffer
   Position = 0
   mData.mvardata = FixedString(-1)
   
   mMemOnlyMode = True

End Sub
Private Sub MemFlush()

   mMemOnlyMode = False

 ' Write buffer back into file
   StorePos
   
   
   Position = 0
   FixedString(-1) = mData.mvardata
   
   SetEndOfFile hFile
   
   
   RestorePos

   
End Sub


Public Sub WriteChanges()
  
  'only working if OnlyMode is on
   Debug.Assert mMemOnlyMode = True

   mMemOnlyMode = False
   MemFlush
   mMemOnlyMode = True
   
End Sub



