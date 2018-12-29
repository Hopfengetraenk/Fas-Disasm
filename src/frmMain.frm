VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   ClientHeight    =   7992
   ClientLeft      =   120
   ClientTop       =   636
   ClientWidth     =   11748
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   7992
   ScaleWidth      =   11748
   Begin VB.CheckBox Chk_Brancher 
      Caption         =   "Disable Brancher"
      Height          =   432
      Left            =   3360
      TabIndex        =   18
      ToolTipText     =   $"frmMain.frx":030A
      Top             =   480
      Visible         =   0   'False
      Width           =   1020
   End
   Begin MSComctlLib.Slider Slider_Zoom 
      Height          =   432
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Zoom Listview"
      Top             =   0
      Width           =   252
      _ExtentX        =   445
      _ExtentY        =   762
      _Version        =   393216
      MousePointer    =   15
      Orientation     =   1
      LargeChange     =   25
      SmallChange     =   25
      Min             =   40
      Max             =   115
      SelStart        =   44
      TickStyle       =   3
      TickFrequency   =   20
      Value           =   44
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   4680
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   6612
      _ExtentX        =   11663
      _ExtentY        =   445
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CheckBox Chk_Cancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   10080
      Style           =   1  'Grafisch
      TabIndex        =   12
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmd_forward 
      Caption         =   "Forward >>>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1590
      TabIndex        =   2
      ToolTipText     =   "Insert or '-'"
      Top             =   60
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'Kein
      Height          =   600
      Left            =   3240
      TabIndex        =   3
      Top             =   0
      Width           =   7572
      Begin VB.CheckBox chk_Inspector 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         Caption         =   "Inspector"
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   3120
         MaskColor       =   &H8000000F&
         Style           =   1  'Grafisch
         TabIndex        =   8
         ToolTipText     =   "Shows Inspector Window"
         Top             =   120
         Value           =   2  'Zwischenzustand
         Width           =   1020
      End
      Begin VB.CheckBox chk_Search 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         Caption         =   "Search"
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   1440
         Style           =   1  'Grafisch
         TabIndex        =   6
         ToolTipText     =   "Search for commands"
         Top             =   120
         Width           =   900
      End
      Begin VB.CheckBox ChkLog 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         Caption         =   "Log"
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   2400
         MaskColor       =   &H80000005&
         Style           =   1  'Grafisch
         TabIndex        =   7
         ToolTipText     =   "Shows Log Window"
         Top             =   120
         Value           =   2  'Zwischenzustand
         Width           =   660
      End
      Begin VB.CheckBox Chk_HexWork 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         Caption         =   "HexWorkShop"
         ForeColor       =   &H80000008&
         Height          =   312
         Left            =   4200
         Style           =   1  'Grafisch
         TabIndex        =   9
         ToolTipText     =   "Opens HexWorkshop when you select a FAS command"
         Top             =   120
         Width           =   1260
      End
      Begin VB.CheckBox chk_Progressbar 
         Caption         =   "Progressbar"
         Height          =   195
         Left            =   5520
         TabIndex        =   5
         ToolTipText     =   "Disable to speed up decrypting"
         Top             =   0
         Value           =   1  'Aktiviert
         Width           =   1260
      End
      Begin VB.CheckBox chk_verbose 
         Caption         =   "Verbose"
         Height          =   195
         Left            =   5520
         TabIndex        =   11
         ToolTipText     =   "Disable to speed up decrypting"
         Top             =   240
         Width           =   1020
      End
      Begin VB.CheckBox chk_Decryptonly 
         Caption         =   "Decrypt only"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   $"frmMain.frx":03C5
         Top             =   0
         Width           =   1260
      End
      Begin VB.CheckBox Chk_cleanup 
         Caption         =   "CleanUp"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Deletes temporary files (*.fct; *.res; *.key)"
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   1020
      End
   End
   Begin VB.Timer Timer_Winhex 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9240
      Top             =   120
   End
   Begin VB.Timer Timer_DropStart 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11040
      Top             =   120
   End
   Begin VB.CommandButton cmd_back 
      Caption         =   "Back <<<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   30
      TabIndex        =   1
      ToolTipText     =   "Backspace or '-'"
      Top             =   60
      Width           =   1575
   End
   Begin MSComctlLib.ListView LV_Log 
      Height          =   6492
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   11292
      _ExtentX        =   19918
      _ExtentY        =   11451
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Liberation Sans"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "id"
         Text            =   "#"
         Object.Width           =   353
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "pos"
         Text            =   "Pos"
         Object.Width           =   1358
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "cmd"
         Text            =   "Cmd"
         Object.Width           =   741
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "disasm"
         Text            =   "Disasm"
         Object.Width           =   1729
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   "params"
         Text            =   "Params"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   "sp"
         Text            =   "SP"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "descr"
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "decomp"
         Text            =   "Decompiled"
         Object.Width           =   7056
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Unten ausrichten
      Height          =   348
      Left            =   0
      TabIndex        =   17
      Top             =   7644
      Width           =   11748
      _ExtentX        =   20722
      _ExtentY        =   614
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3535
            MinWidth        =   3528
            Object.ToolTipText     =   "Current File"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3535
            MinWidth        =   3528
            Object.ToolTipText     =   "Stage of processing"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13117
            MinWidth        =   3528
            Object.ToolTipText     =   "Details"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Gulim"
         Size            =   9.6
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   6432
      Left            =   0
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manuell
      TabIndex        =   16
      Text            =   "frmMain.frx":0452
      Top             =   960
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
      TabIndex        =   13
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

Private Const log_parameter_Show_size As Boolean = 0

#If DoDebug = 0 Then
   Private Const InterpretingProgress_FORMUPDATE_EVERY& = 300
#Else
   Private Const InterpretingProgress_FORMUPDATE_EVERY& = 100
#End If


Private Winhex As New clsSendToWinhex


Const WM_CHAR = &H102
'Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public break As Boolean

Private nav_PositionHistory As New Stack
Private nav_TopStack As New Stack



Public WithEvents File As FasFile
Attribute File.VB_VarHelpID = -1
Private FilePath

Private FileNr As Integer     'Shows Actual FileListIndex
   
Private frmWidth, frmheight As Long


 
Public LispFileData As New Stack
 
'GUI StatusBar Panels Names
Enum Panel
   PanelFilename = 1
   PanelStatus = 2
   PanelDetails = 3
End Enum

Public LV_Log_Ext As New Ext_ListView



Private Sub StartWork()
   
   On Error GoTo StartWork_err
   FileNr = 1
   
   Chk_Cancel.value = False
   Chk_Cancel.Visible = True
   
  ' Winhex.CloseHexWorkshop
   
   Dim item, i&: i = 1
 ' Note:This customized 'For each' is need because filelist may change inside loop
   Do While i <= Filelist.count
          item = Filelist(i)
   
'   For item = LBound(Filelist) - (FilePath <> Empty) To UBound(Filelist)
           Panel_File = item
           Panel_File_ToolTip = Right(FilePath, 50)
           
            'Panel_Status =  tmp, 1
            Filename = FilePath & item 'Filelist(item)
            AddtoLog "Opening File " & Filename
         
      Set File = New FasFile
      
'      File.Create (IIf(FilePath = Empty, Filelist(item),
 '                                      FilePath & "\" & Filelist(item)))
      
      Dim isLsp As Boolean
      isLsp = LspFile_Decrypt(Filename)
      
      If isLsp = False Then
      
       ' output file
         FileLog_open

         
       ' Start Decompiling...
         On Error Resume Next
         File.create Filename

      End If
      
            If Chk_Cancel Or _
               (Err = ERR_GUI_CANCEL) Then
               Dim tmp$
               tmp = "Batch processing canceled !"
               Panel_Status = tmp: AddtoLog tmp
               
               Chk_Cancel.Enabled = True
               Exit Do
            End If


      
    '  Set File = Nothing
      FileNr = FileNr + 1
   i = i + 1: Loop
   
   Chk_Cancel.Visible = False

StartWork_err:
   Select Case Err.Number
   
   Case 0
   
   Case ERR_GUI_CANCEL:
      Chk_Cancel.Enabled = True
   
   Case Is < 0 'Object orientated Error
      AddtoLog "ERROR: " & Err.Description
      Panel_Detail = "ERROR: " & Err.Description
      Resume Next
      
   Case Else:
      MsgBox Err.Number & ": " & Err.Description, vbCritical, "Unexpected Runtime Error"
      Resume Next
   End Select
   
 ' Clear Filelist
   Set Filelist = New Collection
   
Finally:
   
'   Set File = Nothing
   FileLog_close

End Sub


Private Sub Chk_Brancher_Click()
'   File.Cond_Disable = Chk_Brancher.value = vbChecked
End Sub

Private Sub Chk_Cancel_Click()
   If Chk_Cancel = vbChecked Then
      AddtoLog "Cancel request by user"
      Chk_Cancel.Enabled = False
   End If
End Sub


Private Sub Chk_HexWork_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbRightButton Then
      LV_Log.HoverSelection = Not LV_Log.HoverSelection
      Chk_HexWork.FontBold = LV_Log.HoverSelection
   End If
End Sub

Private Sub chk_Inspector_Click()
   frmInspector.Visible = chk_Inspector.value = vbChecked
End Sub

Private Sub chk_Search_Click()
   FrmSearch.Visible = chk_Search.value = vbChecked
End Sub

Private Sub ChkLog_Click()
   frmlog.Visible = ChkLog
End Sub



Private Sub File_initBegin()
   ProgressBar1.Visible = False
   AddtoLog "Initialising ..."
   Panel_Status = "Analysing Data..."
End Sub

Private Sub File_DecryptingBegin(BytesToProgress As Long)
         
         Panel_Status = "Decrypting Data..."
         AddtoLog "Decrypting ..."
         
         ProgressBar1.Min = 0
         ProgressBar1.value = 0
         ProgressBar1.Max = BytesToProgress
         ProgressBar1.Visible = True
         Text1 = Empty
         LV_Log.Visible = False
'
'         Panel_Status =  ("No Valid FSL-File !")
End Sub
Private Sub File_DecryptingProgress(BytesProgressed As Long, CharDecrypted As Long)
   If chk_verbose = vbChecked Then
      PostMessage Text1.Hwnd, WM_CHAR, CharDecrypted And (CharDecrypted > 32), 0
   End If
   
   
   Gui_CheckforCancel
   
   If Chk_Cancel Then
      Dim tmp$
      tmp = "Decrypting canceled !"
      Panel_Status = tmp: AddtoLog tmp
      Err.Raise vbObjectError, "", tmp
   End If
   
   ProgressBar_Update BytesProgressed
   
End Sub

Private Sub ProgressBar_Update(NewValue)

   Static count
   Inc count
   If count > (InterpretingProgress_FORMUPDATE_EVERY * 10) Then
      count = 0
      
      If chk_Progressbar = vbChecked Then ProgressBar1 = NewValue
      
      DoEvents
      
   End If

End Sub


Private Sub File_DecryptingDone()
'         Panel_Status =  IIf(IsDecryptingDone, "Done !", "Nothing done. File is already decrypted !")
         Panel_Status = "Decrypting done !"
         AddtoLog ("Decrypting done !")
End Sub

Private Sub File_InitDone()
   Panel_Status = "Init done !"
End Sub

'
Private Sub File_InterpretingBegin(BytesToProgress As Long)
   Panel_Status = ("Interpreting Data...")
   AddtoLog ("Interpreting Data...")
   LV_Log.Visible = True
   
   ProgressBar1.Min = 0
   ProgressBar1.value = 0
   ProgressBar1.Max = Max(BytesToProgress, 1)
   ProgressBar1.Visible = True
   
End Sub

Private Sub File_InterpretingDone()
         Panel_Status = "Disassembling done !"
         AddtoLog ("Disassembling done !")
         ProgressBar1.Visible = False
         
       ' Jump to DefunMain Function
         Dim item As ListItem
         Set item = LV_Log_Ext.OffsetKeyGet(0, OffToStr(File.ResourceStream_DefunMain))
         LV_Log_Ext.EnsureVisible item
         item.Selected = True
         
         
End Sub

Private Sub Listview_OutputLine(outp As Log_OutputLine, LineBreaksCount, FasCmdlineObj)
   
   With outp
   
      Dim li As MSComctlLib.ListItem
      Set li = LV_Log.ListItems.add(, , LV_Log.ListItems.count)
       
    ' Bind FasCmdlineObj to Listitem
      Set li.Tag = FasCmdlineObj
           
     On Error Resume Next
      
    'Add Key for jump to offset
     LV_Log_Ext.OffsetKey(li, FasCmdlineObj.ModulId) = OffToStr(FasCmdlineObj.Position)
     
     On Error GoTo 0
     Dim TextColor&
     TextColor = GetColor_Cmd("&h" & outp.Command_Byte)

   ' #1 Offset
     Dim LSU As ListSubItem
     Set LSU = LV_Log_Ext.ListSubItem(li, "pos")
     LSU = .offset
     LSU.ForeColor = IIf(File.m_MVar, _
                        ColorConstants.vbMagenta, _
                        ColorConstants.vbBlue _
                        )
   
   ' #2 Command
     Dim LVSI As ListSubItem
     Set LVSI = LV_Log_Ext.ListSubItem(li, "cmd")
     With LVSI
         .Bold = True
         .ForeColor = TextColor
         .Text = outp.Command_Byte
         .ToolTipText = outp.DisASM
     End With
     
   ' #3 Parameters
     With LV_Log_Ext.ListSubItem(li, "params")
         .ForeColor = TextColor
         .Text = outp.Params_Bytes
     End With

   ' #4 Parameters
     With LV_Log_Ext.ListSubItem(li, "disasm")
         .ForeColor = TextColor
         .Text = outp.DisASM
     End With
     
     LV_Log_Ext.ListSubItem(li, "sp") = .Stack
     
   ' #5 Description & DeCompiled
     LV_Log_Ext.ListSubItem(li, "descr") = .Description
     LV_Log_Ext.ListSubItem(li, "decomp") = .DeCompiled

    ' Add linebreaks
    '  Dim LineBreaksCount
    '  Output_GetLineBreaks .Description, LineBreaksCount

      For i = 1 To Min(LineBreaksCount, 2)
         LV_Log.ListItems.add
      Next
      
      Listview_ScrollToItem li

   End With
End Sub
   




Private Sub Gui_CheckforCancel()
   If Chk_Cancel Then
      Dim tmp$
      tmp = "Interpreting canceled !!!"
      
      Panel_Status = tmp: AddtoLog tmp
      FileLog_Add tmp
      
      
      LV_Log.ListItems.add , , tmp
      
      FrmMain.LispFileData.push tmp

     Exit Sub
      
   End If
   
End Sub

Private Sub File_InterpretingProgress(FasCmdlineObj As FasCommando)
   
   Gui_CheckforCancel
   
   ProgressBar_Update FasCmdlineObj.Position
   
   
   Dim Out As Log_OutputLine
 
 ' Format Offset
'   tmp = Format(FasCmdlineObj.Position, "00000")
   Out.offset = BlockAlign_r(OffToStr(FasCmdlineObj.Position), 6)

   
 ' Command_Byte
   Out.Command_Byte = Hex(FasCmdlineObj.Commando)

  
 ' Parameters
   Dim Params, item
   ReDim Params(FasCmdlineObj.Parameters.count)
   
   On Error Resume Next
   Dim i
   For i = 1 To FasCmdlineObj.Parameters.count
      Let item = FasCmdlineObj.Parameters(i)
      Set item = FasCmdlineObj.Parameters(i)
     
      If log_parameter_Show_size And _
         TypeOf item Is T_INT Then
         Params(i) = Format(item, String(item.size * 2, "0"))
      Else
         Params(i) = item
      End If
   Next
   
   'On Error Resume Next
   Out.Params_Bytes = Join(Params)
   Err.Clear   ' Assume: .Params_Bytes are initialised with ""
   
  
   Out.DisASM = FasCmdlineObj.Disassembled_Short
   
   Out.Description = FasCmdlineObj.Disassembled
  
   
 ' Stack
   Out.Stack = File.FasStack.esp
   
 ' Decompiled
   Out.DeCompiled = FasCmdlineObj.Interpreted
    
   Output_DecompiledLine Out.DeCompiled
   
   
   
   Dim LineBreaksCount
   Output_GetLineBreaks Out.Description, LineBreaksCount
   
   DoLog_OutputLine Out, LineBreaksCount


  'Omit listview output to speed up decompiling and reduce memory footprint
   If chk_Decryptonly = vbUnchecked Then
      Listview_OutputLine Out, LineBreaksCount, FasCmdlineObj
   End If
   
   asd
   
End Sub


Sub Output_DecompiledLine(TextLine)
 ' Write Output to *.lsp
   If Trim(TextLine) <> "" Then
      
      LispFileData.push TextLine

      'Print #2, FasCmdlineObj.Interpreted
   End If

End Sub
   
   
Private Sub Listview_ScrollToItem(li As MSComctlLib.ListItem)

   Static count
   count = count + 1
   If (count > InterpretingProgress_FORMUPDATE_EVERY) Then
      count = 0
      If (chk_verbose.value = vbChecked) Then
         LV_Log_Ext.EnsureVisible li
      End If
      
      DoEvents
   End If
   
End Sub

Private Sub Form_Initialize()
Dim a
a = And_(True)


 ' Bind Listview extender to listview
   LV_Log_Ext.create LV_Log
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   break = True
End Sub
Sub LV_Log_ColumnHeadersSize_restore()

  'Restore Listview Columns
   Dim CH As ColumnHeader, tmp$
   For Each CH In LV_Log.ColumnHeaders
      CH.Width = GetSetting(App.EXEName, "Listview", CH.Key, CH.Width)
'      Debug.Print CH.Width
   Next

End Sub
Sub LV_Log_ColumnHeadersSize_save()
   Dim CH As MSComctlLib.ColumnHeader, tmp$
   For Each CH In LV_Log.ColumnHeaders
      SaveSetting App.EXEName, "Listview", CH.Key, CH.Width
   Next
End Sub

Private Sub Form_Load()


   frmWidth = Me.Width
   frmheight = Me.Height
   Me.Caption = App.Title & " V " & App.Major & "." & App.Minor
   Me.Visible = True
   
   On Error GoTo Form_Load_err
   
   LV_Log_ColumnHeadersSize_restore
   
'Test for Commandline Arguments
   Dim CommandLine As New CommandLine
   If CommandLine.NumberOfCommandLineArgs <= 0 Then
      
      mi_open_Click
  
   Else
      Dim item
      For Each item In CommandLine.getArgs()
         Dim dummy As New ClsFilename
         dummy = item
         Filelist.add dummy.Name & dummy.Ext
      Next
      FilePath = dummy.Path
      Call StartWork
   End If
   
   
   FormSettings_Load Me
   
   
   On Error GoTo Form_Load_err
   Exit Sub
Form_Load_err:
   MsgBox Err.Number & ": " & Err.Description, vbCritical, "Runtime Error"
End Sub

Property Let Panel_File_ToolTip(Text$)
   SetBarToolTipText Text, _
               Panel.PanelFilename
End Property

Property Let Panel_File(Text$)
   SetBarText Join(Array("File", FileNr, Text), " "), _
               Panel.PanelFilename
End Property

Property Let Panel_Status(Text$)
   SetBarText Text, Panel.PanelStatus
End Property

Property Let Panel_Detail(Text$)
   SetBarText Text, Panel.PanelDetails
End Property


Private Sub SetBarText(Text$, Optional Panelidx = Panel.PanelStatus)
   StatusBar1.Panels(Panelidx).Text = Text
End Sub


Private Sub SetBarToolTipText(Text$, Optional Panelidx = Panel.PanelStatus)
 'Attention Tooltip is somehow cropped after 50 chars
   StatusBar1.Panels(Panelidx).ToolTipText = Text
End Sub

Public Sub AddtoLog(TextLine$)

   FrmMain.Panel_Detail = TextLine
   
   Dim item
   For Each item In Split(TextLine, vbCrLf)
      frmlog.listLog.AddItem item
   Next
   
End Sub



Private Sub Form_Resize()
   On Error Resume Next
   Dim item ' As Panel
   Dim frmScaleWidth, frmScaleheight As Single
       frmScaleWidth = Me.Width / frmWidth
       frmScaleheight = Me.Height / frmheight
       
'   For Each item In StatusBar1.Panels
'      item.Width = frmScaleWidth * item.Width
'   Next
   frmWidth = Me.Width
   frmheight = Me.Height
   
   Text1.Width = Me.Width - 250
   Text1.Height = Me.Height - Text1.Top - StatusBar1.Height - 800
   
   
   LV_Log.Width = Text1.Width
   LV_Log.Height = Text1.Height
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
  
  
   Dim Form
   For Each Form In Forms
      Unload Form
   Next
   
   FormSettings_Save Me, "LV_Log Text1"
   
   'End 'unload all otherforms
End Sub



'Private Sub LV_Log_jump(item As MSComctlLib.ListItem)
'Private Sub LV_Log_jump(index&)

'End Sub

Private Sub cmd_forward_Click()
   On Error Resume Next
   Nav_forward
End Sub

Private Sub cmd_back_Click()
   On Error Resume Next
   Nav_back
End Sub

Private Sub Inspector_update()

  If frmInspector.Visible = False Then Exit Sub
  
  On Error Resume Next
  frmInspector.updateData LV_Log.SelectedItem.Tag
  If Err Then frmInspector.clean
  
End Sub


Private Sub frmInspector_clean()
   With frmInspector
      .ModulId = ""
      .Position = ""
      .Commando = ""
      .Parameters = ""
      .Disassembled = ""
      .Interpreted = ""
      .Stack_Pointer = ""
      .Stack = ""
      
   End With

End Sub


Private Sub frmInspector_updateData()

   Dim item As FasCommando
   Set item = LV_Log.SelectedItem.Tag
   
   Dim Dest As frmInspector
   Set Dest = frmInspector
   
   With frmInspector
      .ModulId = item.ModulId
      .Position = item.Position
      .Commando = item.Commando
      Dim tmp As clsStrCat
      .Parameters = CStr(Join(CollectionToArray(item.Parameters)))
      .Disassembled = item.Disassembled
      .Interpreted = item.Interpreted
      .Stack_Pointer = item.Stack_Pointer_After
      .Stack = item.Stack_After
      
   End With
   

End Sub



Private Sub LV_Log_DblClick()
   Nav_to
End Sub

Private Sub Nav_forward()
   Dim item As MSComctlLib.ListItem
   
   If nav_PositionHistory.esp < nav_TopStack Then
      
      nav_PositionHistory.esp = nav_PositionHistory.esp + 1
      
      cmd_forward.Enabled = nav_PositionHistory.esp < nav_TopStack
      cmd_back.Enabled = True
      
      LV_Log.SelectedItem.Selected = False
      LV_Log.SelectedItem.Bold = True
      
      'note that the stackpointer + 2 (<= +1 +1)
      Set item = nav_PositionHistory.Storage(nav_PositionHistory.esp + 1)
      With item
       ' Jump to target
         .Bold = False
          
         .Selected = True
         LV_Log_Ext.EnsureVisible item

      End With
            
      LV_Log.SetFocus
      
   End If
End Sub

Private Sub Nav_back()
On Error Resume Next
   Dim item As MSComctlLib.ListItem
   
   If nav_PositionHistory.esp Then
      
      nav_PositionHistory.popIntoVoid
      
      cmd_back.Enabled = nav_PositionHistory.esp
      cmd_forward.Enabled = True
      
      
      LV_Log.SelectedItem.Selected = False
      
      Set item = nav_PositionHistory.Storage(nav_PositionHistory.esp + 1)
      With item
       ' Jump to target
         .Bold = False
          
         .Selected = True
         LV_Log_Ext.EnsureVisible item

      End With
      
      LV_Log.SetFocus
      
   End If
End Sub

Private Sub Nav_to()
   On Error GoTo ERR_Nav_to


   Dim item As MSComctlLib.ListItem
   'Scan selected line of dasm for offset
   ' well numbers will only be found if there is a space before like
   ' "goto $12AF" or " at $0730"
   ' but not "0x222", " 131" or "else jump121"
   ' split line at spaces
   
   Dim RawTextPart
   Set item = LV_Log.SelectedItem
   
   Dim FasCommando As FasCommando
   Set FasCommando = item.Tag
   
'
   Panel_Status = "Try to quickjump from " & OffToStr(FasCommando.Position) & "_" & FasCommando.ModulId
'   If FasCommando.ModulId <> 1 Then
'      Panel_Detail = "Quickjump only works in the fas-function stream."
'      'Err.Raise vbObjectError, "FrmMain::Nav_to", "Quickjump only works in the fas-function stream."
'      Exit Sub
'   Else
      Panel_Detail = " "
'   End If
   
   Dim Rawtext
   Rawtext = LV_Log_Ext.ListSubItem(item, "descr")
   On Error Resume Next
   For Each RawTextPart In Split(Rawtext)
      
      
    ' try to extract number
      If OffToVal(RawTextPart) <> 0 Then
      
       ' Check for valid offset (can listitem with .key be found)
         On Error Resume Next
         Dim moduleID
         moduleID = IIf(Rawtext Like "*Modul:0*", 0, 1)
         Set item = LV_Log_Ext.OffsetKeyGet(moduleID, RawTextPart)
         If Err = 0 Then
            
          ' store current position on Stack
            nav_PositionHistory.push LV_Log.SelectedItem
            nav_TopStack = nav_PositionHistory.esp
            
          ' store new temporally position on Stack aswell
            nav_PositionHistory.push item
            
 '         ' that's to make it temporarely
            nav_PositionHistory.popIntoVoid
            
           
           'mark current LI and save it (-its  position)
            LV_Log.SelectedItem.Bold = True
'            LV_Log.SelectedItem.Selected = False
            cmd_back.Enabled = True
            cmd_forward.Enabled = False

           
            item.Selected = True
            LV_Log_Ext.EnsureVisible item
            
            LV_Log.SetFocus
            
         End If
         Err.Clear
      End If
   Next
   
Exit Sub
ERR_Nav_to:
   If Err Then Panel_Detail = Err.Description
   
End Sub

Private Sub LV_Log_ItemClick(ByVal item As MSComctlLib.ListItem)
   On Error Resume Next
  
   Inspector_update
  
 ' Filter out empty lines AND Skip if "use HexWorkShop" is unchecked
   If (item = "") Or (Chk_HexWork = vbUnchecked) Then Exit Sub
   'Debug.Assert
   
   
   'Bug: Item with index over 0x1f000 gets 0
   'Set LV_Log.SelectedItem = item
   
   'Reset timer
   Timer_Winhex.Enabled = False
   DoEvents
   Timer_Winhex.Enabled = True

End Sub


Private Sub Slider_Zoom_Scroll()
   On Error Resume Next
   LV_Log.Font.size = Slider_Zoom.value / 10
   
End Sub

Private Sub Timer_Winhex_Timer()
   Timer_Winhex.Enabled = False
   DoEvents

   On Error Resume Next
   Winhex.Winhex_JumpToSelectedItem _
      LV_Log, Me, _
      File.Offset_DataStart, _
      File.Offset_CodeStart
      
   If Err Then
      Panel_Status = "Err_SendHexWorks: " & Err.Description
      Chk_HexWork.value = vbUnchecked

   Else
      Panel_Status = ""

   End If
      
   
End Sub

Private Sub LV_Log_KeyPress(KeyAscii As Integer)
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

   LV_Log_ColumnHeadersSize_save

End Sub

Private Sub mi_reload_Click()
      Dim dummy As New ClsFilename
      dummy = Filename
      Filelist.add dummy.Name & dummy.Ext
      
      FilePath = dummy.Path

      StartWork
End Sub



Private Sub Text1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   DragEvent data
End Sub

Private Sub List1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   DragEvent data
End Sub

Private Sub Form_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
   DragEvent data
End Sub

Private Sub DragEvent(data As DataObject)
On Error GoTo DragEvent_err
'   If Data.GetFormat(vbCFText) Then
'      Stop
'   End If
'   If Data.GetFormat(vbCFMetafile) Then
'      Stop
'   End If
      

'   If Data.GetFormat(vbCFFiles) Then
      
'      ReDim Filelist(data.Files.count - 1)
 '     Dim i As Integer
  '    For i = LBound(Filelist) To UBound(Filelist)
   '      Filelist(i) = data.Files.item(i + 1)
    '  Next
      
      Dim FileCount&
      Dim item, dummy As New ClsFilename
      For Each item In data.Files
         dummy = item
         
         Panel_Status = "File dragged in from: " & dummy.Path
         
         ' GetAttr may raise Error if file was not found
         If GetAttr(item) <> vbDirectory Then
         
            FileNr = FileCount + 1
            
          ' on first file...
            If FileCount = 0 Then
            
             ' ... reset Filelist
               Set Filelist = New Collection
               
               
            End If
            Inc FileCount
            
            Panel_File = dummy.Name

            
            Filelist.add dummy.Name & dummy.Ext

 
         Else
            MsgBox _
               "Getting all the contained files is not supported." & vbCrLf & _
               "" & vbCrLf & _
               "Please go into the folder - press ctrl+A to select the all files and " & vbCrLf & _
               "then drag them all in here again.", _
               vbExclamation, _
               "Whoops you dragged in a whole folder." _

            Exit For
         End If
      Next
      
      FilePath = dummy.Path

      
      If Filelist.count Then _
         Timer_DropStart.Enabled = True
'   End If
   
DragEvent_err:

   If Err Then _
      Panel_Detail = "Error: " & Err.Description
      Panel_File = dummy.Name
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
            Filelist.add item
         Next
        
       ' extract path
         Dim dummy As New ClsFilename
         dummy = Filelist(1)
       
       ' If more than 1 file remove first - path only entry
         If Filelist.count <= 1 Then
            Filelist.add dummy.Name & dummy.Ext
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
   
If Err = 20477 Then
   On Error GoTo mi_open_err
   
   CommonDialog1.Filename = _
      InputBox("Please correct the FileName: ", _
                "Whoops FileDialog says 'Invalid filename'", _
                 CommonDialog1.Filename _
                )
   Resume
End If
If Err <> 32755 Then _
   MsgBox Err.Number & ": " & Err.Description, vbCritical, "Runtime Error"
End Sub

Private Sub Timer_DropStart_Timer()

   Timer_DropStart.Enabled = False
   
   StartWork

End Sub

