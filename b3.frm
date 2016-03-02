VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form b3 
   Caption         =   "上海豪曼制冷空调服务有限公司"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdZuan 
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14790
      TabIndex        =   25
      Top             =   0
      Width           =   435
   End
   Begin VB.TextBox txtJxgj 
      Height          =   855
      Left            =   2160
      TabIndex        =   24
      Text            =   "Text12"
      Top             =   2790
      Width           =   12255
   End
   Begin VB.TextBox txtC3 
      Height          =   390
      Left            =   10680
      TabIndex        =   23
      Text            =   "Text11"
      Top             =   2280
      Width           =   3705
   End
   Begin VB.TextBox txtC2 
      Height          =   390
      Left            =   6120
      TabIndex        =   22
      Text            =   "Text10"
      Top             =   2280
      Width           =   3795
   End
   Begin VB.TextBox txtC1 
      Height          =   390
      Left            =   2130
      TabIndex        =   21
      Text            =   "Text9"
      Top             =   2280
      Width           =   3525
   End
   Begin VB.TextBox txtK3 
      Height          =   345
      Left            =   10680
      TabIndex        =   20
      Text            =   "Text8"
      Top             =   1770
      Width           =   3675
   End
   Begin VB.TextBox txtK2 
      Height          =   345
      Left            =   6150
      TabIndex        =   19
      Text            =   "Text7"
      Top             =   1770
      Width           =   3735
   End
   Begin VB.TextBox txtK1 
      Height          =   345
      Left            =   2160
      TabIndex        =   18
      Text            =   "Text6"
      Top             =   1770
      Width           =   3465
   End
   Begin VB.TextBox txtGjqk 
      Height          =   1035
      Left            =   2700
      TabIndex        =   17
      Text            =   "Text5"
      Top             =   6210
      Width           =   11715
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2295
      Left            =   210
      TabIndex        =   15
      Top             =   3780
      Width           =   14205
      _ExtentX        =   25056
      _ExtentY        =   4048
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtYwy 
      Height          =   345
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   540
      Width           =   1485
   End
   Begin VB.TextBox txtBm 
      Height          =   345
      Left            =   4770
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   540
      Width           =   2295
   End
   Begin VB.TextBox txtZw 
      Height          =   345
      Left            =   8610
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   540
      Width           =   1845
   End
   Begin MSComCtl2.DTPicker txtM 
      Height          =   345
      Left            =   12030
      TabIndex        =   26
      Top             =   540
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
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
      Format          =   56950787
      CurrentDate     =   39415
   End
   Begin VB.Label lblKid 
      Caption         =   "lblKid"
      Height          =   225
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label13 
      Caption         =   "改进情况或改进效果评价（由部门主管填写）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   270
      TabIndex        =   16
      Top             =   6210
      Width           =   2145
   End
   Begin VB.Label Label12 
      Caption         =   "绩效改进思路（部门主客指引）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   240
      TabIndex        =   14
      Top             =   2820
      Width           =   1485
   End
   Begin VB.Label Label11 
      Caption         =   "存在问题"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "考核结果"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   210
      TabIndex        =   12
      Top             =   1770
      Width           =   1005
   End
   Begin VB.Label Label9 
      Caption         =   "工作态度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12030
      TabIndex        =   11
      Top             =   1170
      Width           =   1185
   End
   Begin VB.Label Label8 
      Caption         =   "工作能力"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6630
      TabIndex        =   10
      Top             =   1170
      Width           =   1305
   End
   Begin VB.Label Label7 
      Caption         =   "工作业绩"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3570
      TabIndex        =   9
      Top             =   1170
      Width           =   1005
   End
   Begin VB.Label Label6 
      Caption         =   "考核项目"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   210
      TabIndex        =   8
      Top             =   1170
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "姓名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   210
      TabIndex        =   7
      Top             =   570
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "部门"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   6
      Top             =   570
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "岗位"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7560
      TabIndex        =   5
      Top             =   570
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "月份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11160
      TabIndex        =   4
      Top             =   570
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "员工绩效改进计划表"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4650
      TabIndex        =   0
      Top             =   0
      Width           =   4005
   End
End
Attribute VB_Name = "b3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdZuan_Click()
If b1.Visible = True Then
    b2.Visible = True
    b1.Visible = False
ElseIf b2.Visible = True Then
    b3.Visible = True
    b2.Visible = False
ElseIf b3.Visible = True Then
    b1.Visible = True
    b3.Visible = False
End If
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
End Sub


Private Sub Text6_Change()

End Sub

