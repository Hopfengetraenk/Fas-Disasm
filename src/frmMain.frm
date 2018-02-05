VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   630
   ClientWidth     =   11355
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7920
   ScaleWidth      =   11355
   Begin VB.Timer Timer_Winhex 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8280
      Top             =   0
   End
   Begin VB.CheckBox Chk_cleanup 
      Caption         =   "CleanUp"
      Height          =   195
      Left            =   9600
      TabIndex        =   13
      ToolTipText     =   "Deletes temporary files (*.fct; *.res; *.key)"
      Top             =   120
      Value           =   1  'Checked
      Width           =   1740
   End
   Begin VB.CheckBox chk_Decryptonly 
      Caption         =   "Decrypt only"
      Height          =   195
      Left            =   6240
      TabIndex        =   11
      ToolTipText     =   "Decrypt Resoures only - Don't interpret file."
      Top             =   90
      Width           =   1260
   End
   Begin VB.CheckBox chk_verbose 
      Caption         =   "Verbose"
      Height          =   195
      Left            =   5160
      TabIndex        =   10
      ToolTipText     =   "Disable to speed up decrypting"
      Top             =   90
      Width           =   900
   End
   Begin VB.CheckBox chk_Progressbar 
      Caption         =   "Progressbar"
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      ToolTipText     =   "Disable to speed up decrypting"
      Top             =   90
      Value           =   1  'Checked
      Width           =   1260
   End
   Begin VB.CheckBox ChkLog 
      Caption         =   "Log"
      Height          =   255
      Left            =   3240
      MaskColor       =   &H8000000F&
      TabIndex        =   8
      ToolTipText     =   "Show Log Window"
      Top             =   90
      Width           =   660
   End
   Begin VB.CheckBox Chk_HexWork 
      Caption         =   "use HexWorkShop"
      Height          =   195
      Left            =   7800
      TabIndex        =   12
      ToolTipText     =   "Opens HexWorkshop when you select a FAS command"
      Top             =   90
      Width           =   1740
   End
   Begin VB.Timer Timer_DropStart 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9720
      Top             =   0
   End
   Begin VB.CommandButton cmd_forward 
      Caption         =   "Forward >>>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1590
      TabIndex        =   7
      ToolTipText     =   "Insert or '-'"
      Top             =   60
      Width           =   1575
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "Back <<<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   30
      TabIndex        =   6
      ToolTipText     =   "Backspace or '-'"
      Top             =   60
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6975
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12303
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Position"
         Object.Width           =   1111
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Command"
         Object.Width           =   741
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Parameter"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Disassembler"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ESP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Decompiled"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.CheckBox Chk_Cancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   7080
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   7575
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11359
            MinWidth        =   11359
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   6915
      Left            =   0
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "frmMain.frx":030A
      Top             =   480
      Width           =   11175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Menu mi_open 
      Caption         =   "Open"
   End
   Begin VB.Menu mi_reload 
      Caption         =   "Reload"
   End
   Begin VB.Menu mi_Search 
      Caption         =   "Search"
   End
   Begin VB.Menu mi_ColSave 
      Caption         =   "Save ListColumsWidth"
   End
   Begin VB.Menu mi_about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const InterpretingProgress_FORMUPDATE_EVERY& = 300
Const NAV_SCROLLDOWN_LINES& = 10

Const TXTOUT_OPCODE_COL& = 25
Const TXTOUT_DISAM_COL& = 65



Const WM_CHAR = &H102
'Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public break As Boolean

Private nav_PositionHistory As New Stack
Private nav_TopStack As New Stack



Private WithEvents File As FasFile
Attribute File.VB_VarHelpID = -1
Private FilePath

Private FileNr As Integer     'Shows Actual FileListIndex
Private Filename$

Private frmWidth, frmheight As Long


Private ProcID&                'ProcessID of HWORKS32.exe
Private HWORKS_Path
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbSize As Long) As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

 
Private Type KEYBDINPUT
   wVk As Integer
   wScan As Integer
   dwFlags As Long
   time As Long
   dwExtraInfo As Long
End Type
 
 
Private Type INPUT_TYPE
   dwType As Long
   xi As KEYBDINPUT
End Type
 
 
 
'KEYBDINPUT dwFlags-Konstanten
Private Const KEYEVENTF_EXTENDEDKEY = &H1 'Der Scancode hat das Präfix &HE0
Private Const KEYEVENTF_KEYUP = &H2 'Die angegebene Taste wird losgelassen
Private Const KEYEVENTF_UNICODE = &H4 'Benutzt ein Unicode Buchstaben der nicht von einen der Tastaturcodes stammt welcher eine Tastatureingabe Simuliert
 
'INPUT_TYPE dwType-Konstanten
Private Const INPUT_MOUSE = 0 'Mauseingabe
Private Const INPUT_KEYBOARD = 1 'Tastatureingabe
Private Const INPUT_HARDWARE = 2 'Hardwarenachricht
 
 
Public LispFileData As New Stack
 
 
Private Sub SendKey(char As Byte, Optional Keyup = True)
   Dim IT As INPUT_TYPE
   
   With IT
      'KeyDown
       .dwType = INPUT_KEYBOARD
       .xi.wVk = char
       .xi.wScan = MapVirtualKey(char, 0)
       .xi.dwFlags = 0
       If SendInput(1&, IT, 28&) = 0 Then Debug.Print "Sendingkeys failed."
      
      'KeyUp
       .xi.dwFlags = IIf(Keyup, KEYEVENTF_KEYUP, 0)
       If SendInput(1&, IT, 28&) = 0 Then Debug.Print "Sendingkeys failed."
    End With
   
    Sleep (10)
   'Process Windows messages
'    DoEvents
      
End Sub

Private Sub SendKeys(vbKeys$)
   Dim i As Integer
   For i = 1 To Len(vbKeys)
      SendKey (Asc(Mid(vbKeys, i)))
   Next

End Sub




Private Sub StartWork()
   
   On Error GoTo StartWork_err
   FileNr = 1
   
   Chk_Cancel.value = False
   Chk_Cancel.Visible = True
   
   CloseHexWorkshop
   
   Dim item, i&: i = 1
 ' Note:This customized 'For each' is need because filelist may change inside loop
   Do While i <= Filelist.count: item = Filelist(i)
   
'   For item = LBound(Filelist) - (FilePath <> Empty) To UBound(Filelist)
           SetStatus CStr(item), 1
            'SetStatus tmp, 1
            Filename = FilePath & item 'Filelist(item)
            AddtoLog "Opening File " & Filename
         
      Set File = New FasFile
      
'      File.Create (IIf(FilePath = Empty, Filelist(item),
 '                                      FilePath & "\" & Filelist(item)))
      
      Dim isLsp As Boolean
      isLsp = LspFile_Decrypt(Filename)
      
      If isLsp = False Then
      
       ' output file
         Close #1
         Open Filename & ".txt" For Output As 1
         
         
       ' Start Decompiling...
         On Error Resume Next
         File.Create Filename
         

      End If
      
            If Chk_Cancel Or _
               (Err = ERR_GUI_CANCEL) Then
               Dim tmp$
               tmp = "Batch processing canceled !"
               SetStatus tmp: AddtoLog tmp
               Exit Do
            End If


      
    '  Set File = Nothing
      FileNr = FileNr + 1
   i = i + 1: Loop
   
   Chk_Cancel.Visible = False

StartWork_err:
   Select Case Err.Number
   
   Case 0
   
   Case Is < 0 'Object orientated Error
      AddtoLog "ERROR: " & Err.Description
      SetStatus Err.Description
      Resume Next
      
   Case Else:
      MsgBox Err.Number & ": " & Err.Description, vbCritical, "Unexpected Runtime Error"
      Resume Next
   End Select
   
 ' Clear Filelist
   Set Filelist = New Collection

End Sub


Private Sub Chk_Cancel_Click()
   If Chk_Cancel = vbChecked Then
      AddtoLog "Cancel request by user"
      Chk_Cancel.Enabled = False
   End If
End Sub


Private Sub ChkLog_Click()
   frmlog.Visible = ChkLog
End Sub



Private Sub File_initBegin()
   ProgressBar1.Visible = False
   AddtoLog "Initialising ..."
   SetStatus "Analysing Data..."
End Sub

Private Sub File_DecryptingBegin(BytesToProgress As Long)
         
         SetStatus "Decrypting Data..."
         AddtoLog "Decrypting ..."
         
         ProgressBar1.Min = 0
         ProgressBar1.value = 0
         ProgressBar1.Max = BytesToProgress
         ProgressBar1.Visible = True
         Text1 = Empty
         ListView1.Visible = False
'
'         SetStatus ("No Valid FSL-File !")
End Sub
Private Sub File_DecryptingProgress(BytesProgressed As Long, CharDecrypted As Long)
   If chk_verbose = vbChecked Then
      PostMessage Text1.Hwnd, WM_CHAR, CharDecrypted And (CharDecrypted > 32), 0
   End If
   
   If Chk_Cancel Then
      Dim tmp$
      tmp = "Decrypting canceled !"
      SetStatus tmp: AddtoLog tmp
      Err.Raise vbObjectError, "", tmp
   End If
   
   
   Static count
   count = count + 1
   If count > InterpretingProgress_FORMUPDATE_EVERY& Then
      count = 0
      
      If chk_Progressbar = vbChecked Then ProgressBar1 = BytesProgressed
      
      DoEvents
      
   End If

End Sub


Private Sub File_DecryptingDone()
'         SetStatus IIf(IsDecryptingDone, "Done !", "Nothing done. File is already decrypted !")
         SetStatus "Decrypting done !"
         AddtoLog ("Decrypting done !")
End Sub

Private Sub File_InitDone()
   SetStatus "Init done !"
End Sub

Private Sub File_InterpretingBegin()
   SetStatus ("Interpreting Data...")
   AddtoLog ("Interpreting Data...")
   ListView1.Visible = True
End Sub

Private Sub File_InterpretingDone()
         SetStatus "Disassembling done !"
         AddtoLog ("Disassembling done !")
End Sub

Private Sub File_InterpretingProgress(FasCmdlineObj As FasCommando)

   If Chk_Cancel Then
      Dim tmp$
      tmp = "Interpreting canceled !!!"
      SetStatus tmp: AddtoLog tmp
      
      ListView1.ListItems.Add , , tmp
      
      Print #1, tmp

      FrmMain.LispFileData.push tmp

     Exit Sub
      
   End If
   

' Fast and dirty code....
   Dim OutputLine As New clsStrCat ', tmp$
   Dim li As MSComctlLib.ListItem
   
 ' Format Offset
   tmp = Format(FasCmdlineObj.Position, "00000")
   
   Set li = ListView1.ListItems.Add(, , tmp)

 ' Bind FasCmdlineObj to Listitem
   FasCmdlineObj.Tag = File.FasStringtable
   Set li.Tag = FasCmdlineObj
   
   On Error Resume Next
 
 ' add key for quickjump only when interpreting Functionstream (FasStringtable = 1)
   If File.FasStringtable = 1 Then
      li.key = "off:" & FasCmdlineObj.Position
   End If
   
   On Error GoTo 0
   
 ' Offset
   OutputLine.Concat tmp & " " & _
            BlockAlign_l(Hex(FasCmdlineObj.Commando), 5)
   
   li.SubItems(1) = Hex(FasCmdlineObj.Commando)
   
  
 ' Parameters
   Dim item
   For Each item In FasCmdlineObj.Parameters
      Select Case TypeName(item)
         
         Case "Long", "Integer", "Byte"
            OutputLine.Concat " " & Hex(item)
            li.SubItems(2) = li.SubItems(2) & Hex(item) & " "
         
         Case "String"
            OutputLine.Concat """" & item & """ "
            li.SubItems(2) = li.SubItems(2) & """" & item & """ "
      
      End Select
   Next
   

  'Disassembled
  Dim disam$
  disam = FasCmdlineObj.Disassembled
  
  Dim lenBefore&
  lenBefore = Len(disam)
  disam = Replace(disam, vbCrLf, "")
  
  Dim LineBreaksCount&
  
  LineBreaksCount = (lenBefore - Len(disam))
  LineBreaksCount = LineBreaksCount \ Len(vbCrLf)
  
' Format OutputLine to fit in colum
  OutputLine.value = IIf(OutputLine.Length <= TXTOUT_OPCODE_COL, _
            BlockAlign_l(OutputLine.value, TXTOUT_OPCODE_COL), _
            OutputLine.value) & " " & BlockAlign_l(disam, TXTOUT_DISAM_COL)
    li.SubItems(3) = disam
    
  OutputLine.value = BlockAlign_l(OutputLine.value, TXTOUT_OPCODE_COL + TXTOUT_DISAM_COL)
   
 ' Stack
   OutputLine.Concat File.FasStack.ESP
   li.SubItems(4) = File.FasStack.ESP
   
 ' Decompiled
   OutputLine.Concat " " & FasCmdlineObj.Interpreted
   li.SubItems(5) = FasCmdlineObj.Interpreted
   
 ' Newline if ESP=0
   Static lastStackitem
   If File.FasStack.ESP < lastStackitem Then
      LineBreaksCount = LineBreaksCount + 1
   Else
   End If
   lastStackitem = File.FasStack.ESP
 
 ' Add linebreaks
   For i = 1 To LineBreaksCount
      ListView1.ListItems.Add
      OutputLine.Concat vbCrLf
   Next
   
   
 ' Output
'   If FasCmdlineObj.Commando = &H35 Then
'      Debug.Print LineOutput
'   End If
   Print #1, OutputLine.value
   
   
 ' Write Output to *.lsp
   If FasCmdlineObj.Interpreted <> "" Then
      
      LispFileData.push FasCmdlineObj.Interpreted

      'Print #2, FasCmdlineObj.Interpreted
   End If
   
   Static count
   count = count + 1
   If count > InterpretingProgress_FORMUPDATE_EVERY& Then
      count = 0
      li.EnsureVisible
      DoEvents
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   break = True
End Sub

Private Sub Form_Load()


   frmWidth = Me.Width
   frmheight = Me.Height
   Me.Caption = App.Title & " V " & App.Major & "." & App.Minor
   Me.Visible = True
   
   On Error GoTo Form_Load_err
   
   Dim CH As ColumnHeader, tmp$
   For Each CH In ListView1.ColumnHeaders
      CH.Width = GetSetting(App.EXEName, "Listview", CH.Index, CH.Width)
   Next
   
   HWORKS_Path = GetSetting(App.EXEName, "Hexworks", "Path", App.Path & "\Hex Workshop\HWORKS32.exe")
   
'Test for Commandline Arguments
   Dim CommandLine As New CommandLine
   If CommandLine.NumberOfCommandLineArgs <= 0 Then
      
      mi_open_Click
  
   Else
      Dim item
      For Each item In CommandLine.getArgs()
         Dim dummy As New ClsFilename
         dummy = item
         Filelist.Add dummy.Name & dummy.Ext
      Next
      FilePath = dummy.Path
      Call StartWork
   End If
   
   
   On Error GoTo Form_Load_err
   Exit Sub
Form_Load_err:
   MsgBox Err.Number & ": " & Err.Description, vbCritical, "Runtime Error"
End Sub
Private Sub SetStatus(StatusText$, Optional panel = 2)
   StatusBar1.Panels(panel).Text = "File " & FileNr & " " & StatusText
End Sub

Public Sub AddtoLog(Textline$)
   frmlog.listLog.AddItem Textline
End Sub


Private Sub Form_Resize()
   On Error Resume Next
   Dim item As panel
   Dim frmScaleWidth, frmScaleheight As Single
       frmScaleWidth = Me.Width / frmWidth
       frmScaleheight = Me.Height / frmheight
   For Each item In StatusBar1.Panels
      item.Width = frmScaleWidth * item.Width
   Next
   frmWidth = Me.Width
   frmheight = Me.Height
   
   Text1.Width = Me.Width - 250
   Text1.Height = Me.Height - Text1.Top - StatusBar1.Height - 800
   
   
   ListView1.Width = Text1.Width
   ListView1.Height = Text1.Height
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
  
   If HWORKS_Path <> "" Then SaveSetting "Fas-Disasm", "Hexworks", "Path", HWORKS_Path
   
   CloseHexWorkshop
  
   Dim Form
   For Each Form In Forms
      Unload Form
   Next
   'End 'unload all otherforms
End Sub


Private Sub CloseHexWorkshop()
   On Error Resume Next
  
  'close open Hexworks
   Shell Environ("systemroot") & "\System32\Taskkill.exe /pid " & ProcID, vbHide

End Sub

'Private Sub ListView1_jump(item As MSComctlLib.ListItem)
'Private Sub ListView1_jump(index&)

'End Sub

Private Sub cmd_forward_Click()
   Nav_forward
End Sub

Private Sub cmd_back_Click()
   Nav_back
End Sub

Private Sub ListView1_DblClick()
   Nav_to
End Sub

Private Sub Nav_forward()
   Dim item As MSComctlLib.ListItem
   
   If nav_PositionHistory.ESP < nav_TopStack Then
      
      nav_PositionHistory.ESP = nav_PositionHistory.ESP + 1
      
      cmd_forward.Enabled = nav_PositionHistory.ESP < nav_TopStack
      cmd_back.Enabled = True
      
      ListView1.SelectedItem.Selected = False
      ListView1.SelectedItem.Bold = True
      
      'note that the stackpointer + 2 (<= +1 +1)
      With nav_PositionHistory.Storage(nav_PositionHistory.ESP + 1)
         'Jump to target
          .Bold = False
          
          ' Jump to target
          ' 1.to the end
            ListView1.ListItems(ListView1.ListItems.count).EnsureVisible
          ' 2.to the item
            .EnsureVisible
          ' 3.Scroll down some lines
            ListView1.ListItems(.Index - NAV_SCROLLDOWN_LINES).EnsureVisible

          
          .Selected = True
      End With
      
      ListView1.SetFocus
      
   End If
End Sub

Private Sub Nav_back()
   Dim item As MSComctlLib.ListItem
   
   If nav_PositionHistory.ESP Then
      
      nav_PositionHistory.pop
      
      cmd_back.Enabled = nav_PositionHistory.ESP
      cmd_forward.Enabled = True
      
      
      ListView1.SelectedItem.Selected = False
      
      With nav_PositionHistory.Storage(nav_PositionHistory.ESP + 1)
         'Jump to target
          .Bold = False
          
          ' Jump to target
          ' 1.to the end
            ListView1.ListItems(ListView1.ListItems.count).EnsureVisible
          ' 2.to the item
            .EnsureVisible
          ' 3.Scroll down some lines
            ListView1.ListItems(.Index - NAV_SCROLLDOWN_LINES).EnsureVisible

          
          .Selected = True
      End With
      
      ListView1.SetFocus
      
   End If
End Sub

Private Sub Nav_to()

   Dim item As MSComctlLib.ListItem
   
   'Scan selected line of dasm for offset
   ' well numbers will only be found if there is a space before like
   ' "goto 4232" or " at 0730"
   ' but not "0x222" or "else jump121"
   ' split line at spaces
   Dim RawTextPart
   For Each RawTextPart In Split(ListView1.SelectedItem.SubItems(3))
    ' try to extract number
      If Val(RawTextPart) <> 0 Then
        'Check for valid offset (can listitem with .key be found)
         On Error Resume Next
         Set item = ListView1.ListItems("off:" & Val(RawTextPart))
         If Err = 0 Then
            
          ' store current position on Stack
            nav_PositionHistory.push ListView1.SelectedItem
            nav_TopStack = nav_PositionHistory.ESP
            
          ' store new temporally position on Stack alswell
            nav_PositionHistory.push item
          ' that's to make it temporally
            nav_PositionHistory.pop
            
           
           'mark current LI and save it (-its  position)
            ListView1.SelectedItem.Bold = True
'            ListView1.SelectedItem.Selected = False
            cmd_back.Enabled = True
            cmd_forward.Enabled = False

           
          ' Jump to target
          ' 1.to the end
            ListView1.ListItems(ListView1.ListItems.count).EnsureVisible
          ' 2.to the item
            item.EnsureVisible
          ' 3.Scroll down some lines
            ListView1.ListItems(item.Index - NAV_SCROLLDOWN_LINES).EnsureVisible

            item.Selected = True
            
            ListView1.SetFocus
            
         End If
         Err.Clear
      End If
   Next
End Sub

Private Sub ListView1_ItemClick(ByVal item As MSComctlLib.ListItem)
   On Error Resume Next
  
 ' Filter out empty lines AND Skip if "use HexWorkShop" is unchecked
   If (item = "") Or (Chk_HexWork = vbUnchecked) Then Exit Sub
   'Debug.Assert
   Set ListView1.SelectedItem = item
   
   'Reset timer
   Timer_Winhex.Enabled = False
   DoEvents
   Timer_Winhex.Enabled = True

End Sub

Private Sub Timer_Winhex_Timer()
   On Error Resume Next
   Timer_Winhex.Enabled = False
   DoEvents
   
   Dim item As MSComctlLib.ListItem
   Set item = ListView1.SelectedItem

 ' test if Hexworks is still running
   Err.Clear
   AppActivate "Hex Workshop"
   If Err <> 0 Then
    
    ' Start Hexworks
      Do
         Err.Clear
         ProcID = Shell(HWORKS_Path & " """ & Filename & """", vbNormalFocus)
         If Err And (Chk_HexWork <> vbUnchecked) Then
            Dim tmpPath$
            tmpPath = InputBox("Please enter path+Filename to HWORKS32.exe", Err.Description, HWORKS_Path)
            If tmpPath = "" Then
               Chk_HexWork.value = vbUnchecked
               Exit Sub
            Else
               HWORKS_Path = tmpPath
            End If
         Else: Exit Do
         End If
      Loop While True
    
    ' Wait till Hexworks is loaded
'      Sleep (500)
   End If

 ' switch to Hexworks
   AppActivate "Hex Workshop"
 
   
   Dim objFasCmd As FasCommando
   Set objFasCmd = ListView1.SelectedItem.Tag
 
 ' calculate absolute offset in fasfile
   Dim Offset
   Offset = item + IIf(objFasCmd.Tag = 0, File.Offset_DataStart, File.Offset_CodeStart)
 
 ' send goto offset X to Hexworks
   VBA.SendKeys "{ESC}{F5}%b%d" & Offset & "~", 1000
   
 ' Calculate length
   Dim nextitem
   nextitem = ListView1.ListItems(item.Index + 1)
   
    ' Test if next item is empty
      Dim Length
      If nextitem = "" Then
         Length = ListView1.ListItems(item.Index + 2) - item
      Else
         Length = nextitem - item
      End If
   
   ' Errorcorrection
      If Length <= 0 Then Length = 1
   
   
  ' Keyinput: Shift Down - for an unknown reason Hex Workshop 4.10 "+{RIGHT}" don't work
    SendKey vbKeyShift, False
  
  ' Keyinput: SHIFT+RIGHT...
  ' For an unknown reason 'SendKey vbKeyShift, False' don't work Hex Workshop 2.54 but
  ' "+{RIGHT}" does
    VBA.SendKeys "+{RIGHT " & Length & "}", 1
  
  ' Keyinput: Shift UP
    SendKey vbKeyShift

 ' Activated Fas-decomp
   Me.SetFocus
   
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
   Select Case KeyAscii
      
      Case vbKeyBack, 45 '<-vbKeySubtract
         Nav_back
      
      Case 108, 43 'vbKeyAdd & ´
         Nav_forward
      
      Case vbKeyReturn, vbKeySpace
         Nav_to

   End Select
   

End Sub

Private Sub mi_ColSave_Click()

Dim CH As MSComctlLib.ColumnHeader, tmp$
   For Each CH In ListView1.ColumnHeaders
      SaveSetting App.EXEName, "Listview", CH.Index, CH.Width
   Next
End Sub

Private Sub mi_reload_Click()
      Dim dummy As New ClsFilename
      dummy = Filename
      Filelist.Add dummy.Name & dummy.Ext
      
      FilePath = dummy.Path

      StartWork
End Sub

Private Sub mi_Search_Click()
   FrmSearch.Show
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   DragEvent Data
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   DragEvent Data
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   DragEvent Data
End Sub

Private Sub DragEvent(Data As DataObject)
   If Data.GetFormat(vbCFFiles) Then
      
'      ReDim Filelist(data.Files.count - 1)
 '     Dim i As Integer
  '    For i = LBound(Filelist) To UBound(Filelist)
   '      Filelist(i) = data.Files.item(i + 1)
    '  Next
      
      Dim item
      For Each item In Data.Files
         Dim dummy As New ClsFilename
         dummy = item
         Filelist.Add dummy.Name & dummy.Ext
      Next
      FilePath = dummy.Path

      Timer_DropStart.Enabled = True
   End If
End Sub


Private Sub mi_about_Click()
   About.Show vbModal
End Sub

Private Sub mi_open_Click()
   On Error GoTo mi_open_err
   
   mi_open.Enabled = False
      With CommonDialog1
         .DialogTitle = "Select one or more files to open"
         .Filter = "Compiled AutoLISP-file (*.fas *.vlx *.fsl)|*.fas;*.fsl;*.vlx|All files(*.*)|*.*"
'         .Filter = "All files(*.*)|*.*"
         .Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly
         .CancelError = True 'Err.Raise 32755
         .MaxFileSize = 1024
         .ShowOpen
         
        'Convert filenames to list
         Dim item
         Set Filelist = Nothing
         For Each item In Split(.Filename, vbNullChar)
            Filelist.Add item
         Next
        
       ' extract path
         Dim dummy As New ClsFilename
         dummy = Filelist(1)
       
       ' If more than 1 file remove first - path only entry
         If Filelist.count <= 1 Then
            Filelist.Add dummy.Name & dummy.Ext
            FilePath = dummy.Path
         Else
            FilePath = Filelist(1) & "\"
         End If
         Filelist.Remove 1
         
      End With
      
      Call StartWork
      
      mi_open.Enabled = True
      
   Exit Sub
mi_open_err:

   mi_open.Enabled = True
   
If Err <> 32755 Then MsgBox Err.Number & ": " & Err.Description, vbCritical, "Runtime Error"
End Sub

Private Sub Timer_DropStart_Timer()

   Timer_DropStart.Enabled = False
   
   StartWork

End Sub

