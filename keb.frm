VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form keb 
   Caption         =   "绩效考核"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdBack 
      Caption         =   "导航"
      Height          =   585
      Left            =   14490
      Picture         =   "keb.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8550
      Width           =   675
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新  建"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   7950
      TabIndex        =   7
      Top             =   7650
      Width           =   5445
   End
   Begin VB.OptionButton Option3 
      Caption         =   "员工绩效改进计划表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   8010
      TabIndex        =   6
      Top             =   5340
      Width           =   5295
   End
   Begin VB.OptionButton Option2 
      Caption         =   "员工月度考核表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   8010
      TabIndex        =   5
      Top             =   3540
      Width           =   5235
   End
   Begin VB.OptionButton Option1 
      Caption         =   "员工月度工作计划表"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   8010
      TabIndex        =   4
      Top             =   1890
      Width           =   5235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打  开"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7770
      TabIndex        =   3
      Top             =   240
      Width           =   2835
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   12780
      TabIndex        =   1
      Top             =   210
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy年MM月"
      Format          =   53149699
      CurrentDate     =   39399
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8325
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   14684
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11100
      TabIndex        =   2
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "keb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
End Sub
