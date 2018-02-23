VERSION 5.00
Begin VB.Form frmInspector 
   BorderStyle     =   5  '?nderbares Werkzeugfenster
   Caption         =   "Inspector Window"
   ClientHeight    =   4248
   ClientLeft      =   12048
   ClientTop       =   360
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4248
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox StackType 
      Appearance      =   0  '2D
      Height          =   288
      Left            =   3240
      TabIndex        =   10
      Text            =   "Text1"
      ToolTipText     =   "Type Stack"
      Top             =   804
      Width           =   972
   End
   Begin VB.TextBox Stack 
      Appearance      =   0  '2D
      Height          =   648
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmInspector.frx":0000
      ToolTipText     =   "Stack"
      Top             =   480
      Width           =   4236
   End
   Begin VB.TextBox Interpreted 
      Appearance      =   0  '2D
      Height          =   648
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmInspector.frx":0006
      ToolTipText     =   "Interpreted"
      Top             =   3480
      Width           =   4236
   End
   Begin VB.TextBox Disassembled 
      Appearance      =   0  '2D
      Height          =   648
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frmInspector.frx":000C
      ToolTipText     =   "Disasm"
      Top             =   2640
      Width           =   4236
   End
   Begin VB.TextBox Parameters 
      Appearance      =   0  '2D
      Height          =   648
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmInspector.frx":0012
      ToolTipText     =   "Command Parameters"
      Top             =   1800
      Width           =   4236
   End
   Begin VB.TextBox Stack_Pointer 
      Appearance      =   0  '2D
      Height          =   372
      Left            =   3600
      TabIndex        =   1
      Text            =   "Text1"
      ToolTipText     =   "Stack Pointer"
      Top             =   0
      Width           =   612
   End
   Begin VB.TextBox Commando 
      Appearance      =   0  '2D
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1440
      TabIndex        =   5
      Text            =   "Text1"
      ToolTipText     =   "Command"
      Top             =   1320
      Width           =   612
   End
   Begin VB.TextBox Position 
      Appearance      =   0  '2D
      Height          =   372
      Left            =   600
      TabIndex        =   4
      Text            =   "Text1"
      ToolTipText     =   "Offset"
      Top             =   1320
      Width           =   732
   End
   Begin VB.TextBox ModulId 
      Appearance      =   0  '2D
      Height          =   372
      Left            =   0
      TabIndex        =   3
      Text            =   "Text1"
      ToolTipText     =   "ModulID"
      Top             =   1320
      Width           =   492
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   492
      Index           =   3
      Left            =   1800
      TabIndex        =   6
      Top             =   1320
      Width           =   2412
   End
   Begin VB.Label Label1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Stack"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   492
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1692
   End
End
Attribute VB_Name = "frmInspector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myFormExtenter As FormExt_MostTop
Private Sub Form_Initialize()
   Set myFormExtenter = New FormExt_MostTop
   myFormExtenter.Create Me
End Sub

Private Sub Form_Load()
   On Error Resume Next
   clean
End Sub






Public Sub clean()
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


Public Sub updateData(item As FasCommando)
  On Error Resume Next
  updateData2 item
  If Err Then clean
End Sub

Private Sub updateData2(item As FasCommando)
   
   Dim Dest As frmInspector
   Set Dest = frmInspector
   
   With frmInspector
      .ModulId = item.ModulId
      .Position = item.Position
      .Commando = item.Commando
      .Commando.ForeColor = GetColor_Cmd(item.Commando)
      
      Dim tmp As clsStrCat
      .Parameters = CStr(Join(CollectionToArray(item.Parameters)))
      .Disassembled = item.Disassembled
      .Interpreted = item.Interpreted
      .Stack_Pointer = item.Stack_Pointer_After
      
      Dim mTypeName
      mTypeName = TypeName(item.Stack_After)
      .StackType = mTypeName
      .StackType.ForeColor = GetColor_Type(mTypeName)

      .Stack = item.Stack_After
      .Stack.ForeColor = .StackType.ForeColor
      
   End With
   

End Sub


