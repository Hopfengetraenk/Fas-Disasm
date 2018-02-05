VERSION 5.00
Begin VB.Form About 
   Caption         =   "About"
   ClientHeight    =   3195
   ClientLeft      =   1635
   ClientTop       =   4230
   ClientWidth     =   4680
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   1260
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "About.frx":030A
      Top             =   1785
      Width           =   3855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "What is this programm good for:"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1560
      Width           =   2235
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "in Nov 2005 ( Minor Update Dez 2013)"
      Height          =   195
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   2745
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Done by"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   600
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   4680
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://acad-fas.tipido.net/"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CW2K"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "label(0)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
   Label1(0).Caption = App.Title & " V " & App.Major & "." & App.Minor & "." & App.Revision
End Sub



Private Sub Label19_Click()
   ShellExecute 0, "open", Label19.Caption, "", "", 0
End Sub

Private Sub Label2_Click()
   ShellExecute 0, "open", "http://lisp-decompiler-project.cjb.net", "", "", 0
End Sub
