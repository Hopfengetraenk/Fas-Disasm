VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Commandline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CommandLine"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetCommandLineRaw Lib "kernel32" Alias "GetCommandLineA" () As String
Private argsWithQuotes$()
Private Args$()
Public NumberOfCommandLineArgs
Public CommandLine As String


'///////////////////////////////////////////////////
'// getArgs - Get CommandLineArguments with Quotes
Public Function getArgs() As String()
   getArgs = Args
End Function

Public Function getArgsWithQuotes()
   getArgsWithQuotes = argsWithQuotes
End Function


'//////////////////////////////////////////////////////////////////////////
'// Constructor - Sets args, args and NumberOfCommandLineArgs
Private Sub Class_Initialize()
   
  GetCommandLine

  If CommandLine = Empty Then Exit Sub
  Call commandLineSplitRek(CommandLine)

End Sub


'////////////////////////////////////////////////////////////////////////////////////////////////
'// commandLineSplitRek -  Split the commandLine and store chunks in args() and argsWithQuotes()
'//                        should only used by the Constructor
Private Sub commandLineSplitRek(CommandLine As String, Optional oldPos% = 1, Optional levelCounter = 0)
   Dim endPos%                ' Endposition for cut (...e.exe"<-)
   Dim newStartPos%           ' Startposition for "new" commandline (->"C:\p...")
   Dim bIsQuoted As Boolean
   Dim bEndLine As Boolean
   Dim bNoQuotedAtEnd As Boolean
   
   ' does Commandline starts with "C:\P...
     bIsQuoted = Left$(CommandLine, 1) = """"
     If bIsQuoted Then
       'find next "
        endPos = InStr(2, CommandLine, """")
     Else
       'deincrement endpos -delete last quote
        endPos = InStr(1, CommandLine, " ") - 1
     End If
     
     bEndLine = endPos = Len(CommandLine)
     bNoQuotedAtEnd = bIsQuoted And bEndLine And (endPos = 0)
          
   ' Check if we reached the end
     If endPos <= 0 Or bEndLine Then
     
      ' we are at the end so endPos = length of CommandLine
        endPos = Len(CommandLine)
        
      ' create array for saving commandline arguments
        NumberOfCommandLineArgs = levelCounter + 1
        ReDim argsWithQuotes(NumberOfCommandLineArgs - 1)
        ReDim Args(NumberOfCommandLineArgs - 1)

     
     Else
     
             
      ' Filter out any whitespaces (for ex. 2 Spaces like ..exe"   "C:\..)
        For newStartPos = endPos + 1 To Len(CommandLine)
            If Mid(CommandLine, newStartPos, 1) > " " Then Exit For
        Next
             
      ' Call commandLineSplitRek recursiv with "new" commandline 6 increase levelCounter
        commandLineSplitRek Mid(CommandLine, newStartPos), newStartPos, levelCounter + 1
     End If
    'Save Data in Array
     argsWithQuotes(levelCounter) = Mid(CommandLine, 1, endPos)
     If bIsQuoted Then
        Args(levelCounter) = Mid(CommandLine, 2, endPos - 2)
     Else
        Args(levelCounter) = Mid(CommandLine, 1, endPos)
     End If
End Sub
Private Function GetCommandLine() As String

   CommandLine = Command
  
End Function

