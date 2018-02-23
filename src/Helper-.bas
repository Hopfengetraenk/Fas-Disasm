Attribute VB_Name = "Help"
Option Explicit

Public Const ERR_FILESTREAM = &H1000000
Public Const ERR_OPENFILE = vbObjectError + ERR_FILESTREAM + 1


Public Const ERR_VLXSPLIT = &H2000000
Public Const ERR_NO_VLX_FILE = vbObjectError + ERR_VLXSPLIT + 1


Public Const ERR_GUIEVENTS = &H3000000
Public Const ERR_GUI_CANCEL = vbObjectError + ERR_GUIEVENTS + 1


Public i, j As Integer

Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10

'Public Filelist
Public Filelist As New Collection

Type Log_OutputLine
   Offset As String
   Command_Byte As String
   Params_Bytes As String
   Description As String
   Stack As String
   DeCompiled As String
End Type


Sub qw()
   FrmMain.break = False
   Do
      DoEvents
   Loop While FrmMain.break = False
End Sub

Sub asd()
   FrmMain.ListView1.ListItems(FrmMain.ListView1.ListItems.count).EnsureVisible
End Sub



Sub MostTop(Hwnd As Long)
   'SetWindowPos Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


Public Function HexvaluesToString$(Hexvalues$)
   Dim tmpchar
   For Each tmpchar In Split(Hexvalues)
      HexvaluesToString = HexvaluesToString & Chr("&h" & tmpchar)
   Next
End Function


Function Max(ParamArray values())
   Dim item
   For Each item In values
      Max = IIf(Max < item, item, Max)
   Next
End Function

Function Min(ParamArray values())
   Dim item
   Min = &H7FFFFFFF
   For Each item In values
      Min = IIf(Min > item, item, Min)
   Next
End Function

Function limit(value, upperLimit, Optional lowerLimit = 0)
   'limit = IIf(Value > upperLimit, upperLimit, IIf(Value < lowerLimit, lowerLimit, Value))

   If (value > upperLimit) Then _
      limit = upperLimit _
   Else _
      If (value < lowerLimit) Then _
         limit = lowerLimit _
      Else _
         limit = value
   
End Function

Function RangeCheck(ByVal value&, Max&, Optional Min& = 0, Optional ErrText, Optional ErrSource$) As Boolean
   RangeCheck = (Min <= value) And (value <= Max)
   If (RangeCheck = False) And (IsMissing(ErrText) = False) Then Err.Raise vbObjectError, ErrSource, ErrText
End Function

Public Function H8(ByVal value As Long)
   H8 = Right("0" & Hex(value), 2)
End Function

Public Function H16(ByVal value As Long)
   H16 = Right(String(3, "0") & Hex(value), 4)
End Function
Public Function H32(ByVal value As Long)
   H32 = Right(String(7, "0") & Hex(value), 8)
End Function

Public Function Dec3$(ByVal value$)
   Dec3 = Right(String(3, "0") & value, 3)
End Function
Public Function Dec2$(ByVal value$)
   Dec2 = Right(String(3, "0") & value, 2)
End Function



Public Function Swap(ByRef a, ByRef b)
   Swap = b
   b = a
   a = Swap
End Function

'////////////////////////////////////////////////////////////////////////
'// BlockAlign_r  -  Erzeugt einen rechtsbündigen BlockString
'//
'// Beispiel1:     BlockAlign_r("Summe",7) -> "  Summe"
'// Beispiel2:     BlockAlign_r("Summe",4) -> "umme"
Public Function BlockAlign_r(RawString, Blocksize) As String
  'String kürzen lang wenn zu
   RawString = Right(RawString, Blocksize)
  'mit Leerzeichen auffüllen
   BlockAlign_r = Space(Blocksize - Len(RawString)) & RawString
End Function

'////////////////////////////////////////////////////////////////////////
'// BlockAlign_l  -  Erzeugt einen linksbündigen BlockString
'//
'// Beispiel1:     BlockAlign_l("Summe",7) -> "Summe  "
'// Beispiel2:     BlockAlign_l("Summe",4) -> "Summ"
Public Function BlockAlign_l(RawString, Blocksize) As String
  'String kürzen lang wenn zu
   RawString = Left(RawString, Blocksize)
  'mit Leerzeichen auffüllen
   BlockAlign_l = RawString & Space(Blocksize - Len(RawString))
End Function

Public Function FileLoad$(Filename$)
   Dim File As New FileStream
   With File
      .Create Filename, False, False, True
      FileLoad = .FixedString(-1)
      .CloseFile
   End With
End Function

Public Sub FileSave(Filename$, data$)
   On Error GoTo err_FileSave
   Dim File As New FileStream
   With File
      .Create Filename, True, False, False
      .FixedString(-1) = data
      .CloseFile
   End With

Exit Sub
err_FileSave:
   Log "ERROR during FileSave: " & Err.Description
End Sub

Public Function Inc(ByRef value, Optional Increment& = 1)
   value = value + Increment
   Inc = value
End Function

Public Function Dec(ByRef value, Optional DeIncrement& = 1)
   value = value - DeIncrement
   Dec = value
End Function
