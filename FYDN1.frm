VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FYDN1 
   Caption         =   "资产、费用申请、报销单"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdD 
      Enabled         =   0   'False
      Height          =   405
      Left            =   14250
      Picture         =   "FYDN1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8760
      Width           =   465
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14760
      Picture         =   "FYDN1.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "返回"
      Top             =   8760
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   13860
      Picture         =   "FYDN1.frx":028C
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "保存"
      Top             =   8790
      Width           =   465
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   13350
      Picture         =   "FYDN1.frx":08F6
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "修改"
      Top             =   8790
      Width           =   465
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   0
      TabIndex        =   4
      Top             =   1050
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "费用申请"
      TabPicture(0)   =   "FYDN1.frx":0C00
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label12"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtgMx"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text4"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "报销"
      TabPicture(1)   =   "FYDN1.frx":0C1C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.TextBox Text4 
         Height          =   585
         Left            =   1350
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   5580
         Width           =   13755
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   10110
         TabIndex        =   18
         Top             =   5220
         Width           =   2085
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   5400
         TabIndex        =   16
         Top             =   5220
         Width           =   2025
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1350
         TabIndex        =   14
         Top             =   5220
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         Caption         =   "审核流程:"
         Height          =   1425
         Left            =   30
         TabIndex        =   12
         Top             =   6210
         Width           =   15165
         Begin VB.CommandButton cmdQm 
            Height          =   345
            Index           =   5
            Left            =   5760
            TabIndex        =   41
            Top             =   510
            Width           =   945
         End
         Begin VB.CommandButton cmdQm 
            Height          =   345
            Index           =   4
            Left            =   4710
            TabIndex        =   38
            Top             =   510
            Width           =   945
         End
         Begin VB.CommandButton cmdQm 
            Height          =   345
            Index           =   3
            Left            =   3690
            TabIndex        =   35
            Top             =   510
            Width           =   945
         End
         Begin VB.CommandButton cmdQm 
            Height          =   345
            Index           =   2
            Left            =   2640
            TabIndex        =   32
            Top             =   510
            Width           =   945
         End
         Begin VB.CommandButton cmdQm 
            Height          =   345
            Index           =   1
            Left            =   1620
            TabIndex        =   29
            Top             =   510
            Width           =   945
         End
         Begin VB.CommandButton cmdQm 
            Height          =   345
            Index           =   0
            Left            =   570
            TabIndex        =   26
            Top             =   510
            Width           =   945
         End
         Begin VB.CommandButton cmdPje 
            Caption         =   "评审建议"
            Height          =   1095
            Left            =   120
            TabIndex        =   25
            Top             =   270
            Width           =   345
         End
         Begin VB.Label lblQM 
            Caption         =   "董事长"
            Height          =   225
            Index           =   5
            Left            =   5850
            TabIndex        =   43
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblTm 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   5
            Left            =   5760
            TabIndex        =   42
            Top             =   930
            Width           =   945
         End
         Begin VB.Label lblQM 
            Caption         =   "财务总监"
            Height          =   225
            Index           =   4
            Left            =   4800
            TabIndex        =   40
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblTm 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   4
            Left            =   4710
            TabIndex        =   39
            Top             =   930
            Width           =   945
         End
         Begin VB.Label lblQM 
            Caption         =   "总经理"
            Height          =   225
            Index           =   3
            Left            =   3780
            TabIndex        =   37
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblTm 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   3
            Left            =   3690
            TabIndex        =   36
            Top             =   930
            Width           =   945
         End
         Begin VB.Label lblQM 
            Caption         =   "财务部"
            Height          =   225
            Index           =   2
            Left            =   2730
            TabIndex        =   34
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblTm 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   2
            Left            =   2640
            TabIndex        =   33
            Top             =   930
            Width           =   945
         End
         Begin VB.Label lblQM 
            Caption         =   "部门经理"
            Height          =   225
            Index           =   1
            Left            =   1710
            TabIndex        =   31
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblTm 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   1
            Left            =   1620
            TabIndex        =   30
            Top             =   930
            Width           =   945
         End
         Begin VB.Label lblTm 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   0
            Left            =   570
            TabIndex        =   28
            Top             =   930
            Width           =   945
         End
         Begin VB.Label lblQM 
            Caption         =   "申请人"
            Height          =   225
            Index           =   0
            Left            =   660
            TabIndex        =   27
            Top             =   240
            Width           =   915
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMx 
         Height          =   4365
         Left            =   0
         TabIndex        =   11
         Top             =   720
         Width           =   15225
         _ExtentX        =   26855
         _ExtentY        =   7699
         _Version        =   393216
         Rows            =   20
         Cols            =   6
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.Label Label12 
         Caption         =   "备注"
         Height          =   555
         Left            =   570
         TabIndex        =   19
         Top             =   5610
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "暂支金额"
         Height          =   225
         Left            =   8760
         TabIndex        =   17
         Top             =   5250
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "批准总额"
         Height          =   255
         Left            =   3750
         TabIndex        =   15
         Top             =   5250
         Width           =   1365
      End
      Begin VB.Label Label9 
         Caption         =   "申请总额"
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   5250
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "区域:"
         Height          =   165
         Left            =   720
         TabIndex        =   10
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label Label4 
         Height          =   195
         Left            =   1980
         TabIndex        =   9
         Top             =   450
         Width           =   2505
      End
      Begin VB.Label Label5 
         Caption         =   "申请部门:"
         Height          =   225
         Left            =   4950
         TabIndex        =   8
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label Label6 
         Height          =   195
         Left            =   6240
         TabIndex        =   7
         Top             =   450
         Width           =   1755
      End
      Begin VB.Label Label7 
         Caption         =   "申请日期:"
         Height          =   225
         Left            =   8670
         TabIndex        =   6
         Top             =   450
         Width           =   1035
      End
      Begin VB.Label Label8 
         Height          =   195
         Left            =   9870
         TabIndex        =   5
         Top             =   450
         Width           =   1845
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "变动费用"
      Height          =   345
      Left            =   6900
      TabIndex        =   3
      Top             =   570
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "固定费用"
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   570
      Width           =   2445
   End
   Begin VB.Label Label2 
      Caption         =   "请选择"
      Height          =   405
      Left            =   870
      TabIndex        =   1
      Top             =   630
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "资产、费用申请、报销单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   90
      Width           =   4035
   End
End
Attribute VB_Name = "FYDN1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
dtgMx.ColWidth(0) = 300
dtgMx.Row = 0
dtgMx.Cols = 7
dtgMx.Col = 1: dtgMx.Text = "费用类别": dtgMx.Col = 2: dtgMx.Text = "内容": dtgMx.Col = 3: dtgMx.Text = "金额"
dtgMx.Col = 4: dtgMx.Text = "项目名称": dtgMx.Col = 5: dtgMx.Text = "合同编号"
dtgMx.ColWidth(1) = 1485: dtgMx.ColWidth(2) = 2220: dtgMx.ColWidth(4) = 2085: dtgMx.ColWidth(5) = 2055: dtgMx.ColWidth(6) = 5600
End Sub
