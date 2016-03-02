VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form htBrow 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查询框"
   ClientHeight    =   9210
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   Begin NiceFormControl.NiceButton OKButton 
      Height          =   405
      Left            =   11310
      TabIndex        =   41
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   714
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "htBrow.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Caption         =   "详     情"
   End
   Begin NiceFormControl.NiceContainr NEDIT 
      Height          =   2625
      Left            =   11190
      TabIndex        =   38
      Top             =   5700
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   4630
      HeaderLightColor=   14078715
      HeaderDarkColor =   11446008
      BackLightColor  =   15790078
      BackDarkColor   =   14736892
      BorderColor     =   10985207
      TextColor       =   1906026
      Caption         =   "新建"
      Theme           =   9
      Begin NiceFormControl.NiceButton NiceButton1 
         Height          =   315
         Left            =   540
         TabIndex        =   44
         Top             =   720
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         BTYPE           =   3
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "htBrow.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Caption         =   "询 价 指 令"
      End
      Begin NiceFormControl.NiceButton cmdHt 
         Height          =   315
         Left            =   540
         TabIndex        =   40
         Top             =   1800
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         BTYPE           =   3
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "htBrow.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Caption         =   "合 同 评 审 单"
      End
      Begin NiceFormControl.NiceButton cmdXjd 
         Height          =   315
         Left            =   540
         TabIndex        =   39
         Top             =   1290
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         BTYPE           =   3
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "htBrow.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Caption         =   "询 价 单"
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   495
      Left            =   10320
      TabIndex        =   31
      Top             =   8730
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   873
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9060
      Top             =   8910
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8430
      Top             =   8820
   End
   Begin VB.Frame frmNew 
      BackColor       =   &H00C0FFC0&
      Caption         =   "制作新流程报价单"
      Height          =   2355
      Left            =   11280
      TabIndex        =   23
      Top             =   5970
      Width           =   3975
      Begin VB.ComboBox txtKhmc 
         Height          =   300
         Left            =   1110
         TabIndex        =   29
         Top             =   870
         Width           =   2325
      End
      Begin MSDataListLib.DataCombo txtXmmc 
         Height          =   330
         Left            =   1080
         TabIndex        =   28
         ToolTipText     =   "请键入关键字后,按回车键,随后在列表中选择项目名称"
         Top             =   330
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00C0FFC0&
         Caption         =   "新建报价单(合同评审单)"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1920
         Width           =   3285
      End
      Begin VB.Label lblTx 
         BackStyle       =   0  'Transparent
         Caption         =   "提示：如果你不能选择签约客户，可能是您未在项目资料中设置客户档案，请查实"
         ForeColor       =   &H000000FF&
         Height          =   585
         Left            =   240
         TabIndex        =   30
         Top             =   1320
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "请先确认项目名称和客户名称"
         ForeColor       =   &H00FF0000&
         Height          =   1575
         Left            =   3480
         TabIndex        =   27
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "项目名称"
         Height          =   225
         Left            =   210
         TabIndex        =   26
         Top             =   390
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "签约客户"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   765
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   7605
      Left            =   11220
      TabIndex        =   9
      Top             =   720
      Width           =   4035
      Begin NiceFormControl.NiceButton cmdWZX 
         Height          =   345
         Left            =   90
         TabIndex        =   32
         Top             =   4050
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   609
         BTYPE           =   3
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "htBrow.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Style           =   9
         Caption         =   "未执行合同"
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "查询方式一"
         Height          =   1545
         Left            =   90
         TabIndex        =   16
         Top             =   30
         Width           =   3945
         Begin NiceFormControl.NiceOption optY 
            Height          =   240
            Index           =   0
            Left            =   270
            TabIndex        =   33
            Top             =   810
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "评审"
            BackColor       =   12648384
         End
         Begin VB.CommandButton cmdRef 
            Caption         =   "刷 新"
            Height          =   735
            Left            =   3450
            TabIndex        =   17
            Top             =   720
            Width           =   375
         End
         Begin VB.Frame frmZZ 
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   15
            Left            =   1290
            TabIndex        =   18
            Top             =   720
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dd2 
            Height          =   285
            Left            =   2550
            TabIndex        =   19
            Top             =   330
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            _Version        =   393216
            CalendarBackColor=   12648384
            CalendarTitleBackColor=   16448
            Format          =   100401153
            CurrentDate     =   38100
         End
         Begin MSComCtl2.DTPicker dd1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dddddd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   690
            TabIndex        =   20
            Top             =   330
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            _Version        =   393216
            CalendarBackColor=   12648384
            CalendarTitleBackColor=   16448
            Format          =   100401153
            CurrentDate     =   38100
         End
         Begin NiceFormControl.NiceOption optY 
            Height          =   240
            Index           =   1
            Left            =   270
            TabIndex        =   34
            Top             =   1200
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "盖章"
            BackColor       =   12648384
         End
         Begin NiceFormControl.NiceOption optY 
            Height          =   240
            Index           =   2
            Left            =   1290
            TabIndex        =   35
            Top             =   810
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "执行"
            BackColor       =   12648384
         End
         Begin NiceFormControl.NiceOption optY 
            Height          =   240
            Index           =   3
            Left            =   1290
            TabIndex        =   36
            Top             =   1200
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "完成"
            BackColor       =   12648384
         End
         Begin NiceFormControl.NiceOption optY 
            Height          =   240
            Index           =   4
            Left            =   2220
            TabIndex        =   37
            Top             =   810
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   423
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "作废"
            BackColor       =   12648384
            SkinIdx         =   20
         End
         Begin VB.Label Label4 
            Caption         =   "截至"
            Height          =   225
            Left            =   2100
            TabIndex        =   22
            Top             =   390
            Width           =   375
         End
         Begin VB.Label lblQRq 
            BackColor       =   &H00C0FFC0&
            Caption         =   "起始"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   390
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "查询方式二"
         Height          =   1905
         Left            =   90
         TabIndex        =   10
         Top             =   1650
         Width           =   3945
         Begin VB.TextBox txtYc 
            Height          =   285
            Left            =   270
            TabIndex        =   13
            Top             =   1380
            Width           =   3015
         End
         Begin VB.ComboBox comXZ 
            Height          =   300
            ItemData        =   "htBrow.frx":008C
            Left            =   270
            List            =   "htBrow.frx":0099
            TabIndex        =   12
            Text            =   "合同编号"
            Top             =   630
            Width           =   3075
         End
         Begin VB.CommandButton cmdRef1 
            Caption         =   "查  询"
            Height          =   825
            Left            =   3450
            TabIndex        =   11
            Top             =   990
            Width           =   345
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "值"
            Height          =   255
            Left            =   300
            TabIndex        =   15
            Top             =   1020
            Width           =   465
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "条件"
            Height          =   255
            Left            =   300
            TabIndex        =   14
            Top             =   300
            Width           =   795
         End
      End
      Begin NiceFormControl.NiceButton command2 
         Height          =   345
         Left            =   90
         TabIndex        =   42
         Top             =   3630
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   609
         BTYPE           =   3
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "htBrow.frx":00BB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Style           =   9
         Caption         =   "全 部"
      End
      Begin NiceFormControl.NiceButton cmdXJView 
         Height          =   345
         Left            =   90
         TabIndex        =   43
         Top             =   4500
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   609
         BTYPE           =   3
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "htBrow.frx":00D7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Style           =   15
         Caption         =   "询价单查询"
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8475
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   14949
      _Version        =   393216
      BackColor       =   12648384
      Rows            =   50
      Cols            =   18
      BackColorFixed  =   16777152
      BackColorBkg    =   -2147483636
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   18
   End
   Begin VB.Frame frmTB 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   8670
      Width           =   7905
      Begin VB.TextBox txtHtze 
         Height          =   285
         Left            =   1380
         TabIndex        =   4
         Top             =   0
         Width           =   1275
      End
      Begin VB.ComboBox comYw 
         Height          =   300
         ItemData        =   "htBrow.frx":00F3
         Left            =   3510
         List            =   "htBrow.frx":00FA
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.ComboBox comKh 
         Height          =   300
         Left            =   5610
         TabIndex        =   2
         Top             =   0
         Width           =   2265
      End
      Begin VB.Label lblHtze 
         BackStyle       =   0  'Transparent
         Caption         =   "按合同金额查询"
         Height          =   225
         Left            =   60
         TabIndex        =   7
         Top             =   60
         Width           =   1275
      End
      Begin VB.Label lblYw 
         BackStyle       =   0  'Transparent
         Caption         =   "业务员"
         Height          =   255
         Left            =   2850
         TabIndex        =   6
         Top             =   60
         Width           =   615
      End
      Begin VB.Label lblKh 
         BackStyle       =   0  'Transparent
         Caption         =   "客户名称"
         Height          =   255
         Left            =   4800
         TabIndex        =   5
         Top             =   60
         Width           =   795
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "退 出"
      Height          =   345
      Left            =   14010
      TabIndex        =   0
      Top             =   8670
      Width           =   1005
   End
End
Attribute VB_Name = "htBrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoBr As Object
Dim ZT As String
Dim adoXmmc As Object
Dim adoKhmc As Object

Dim timZm As Integer '数据提交后,由timWait执行的后续命令ID(1添加新合同,2添加询价指令)
Dim Hid As Long '新添加合同的Hid

Dim PRF As Integer '生成纯配件合同还是其它合同
Dim Ra: Dim La
Public DT As String '查询SQL字段
Dim Bid As Long
'Dim KhId(0 To 6) As Integer
Private Sub CancelButton_Click()
Me.Visible = False
frmZu.Enabled = True
frmZu.SetFocus
End Sub

Private Sub cmdHt_Click()

    Call FMXCXmmc.Qing
    FMXCXmmc.Show
    FMXCXmmc.ZOrder 0
    FMXCXmmc.Lb = "合同评审单"
    FMXCXmmc.NiceButton1.Caption = "生 成 单 据 (合同评审单)"
End Sub

Private Sub cmdNew_Click()
Dim Ti As Integer
timZm = 1
If txtKhmc.Text = "" Or txtKhmc.ToolTipText = "" Then
    MsgBox ("您没有选择正确的客户!")
    Exit Sub
End If
    'MsgBox "0"
    Call modNewHT.NewMQing
       ' MsgBox "1"
''''''PRF = MsgBox("是纯配件，还是其它合同？" & Chr(13) & "选择'是'进行纯配件询价，询价单直接跳至零件事业部" & Chr(13) & "选择'否'则为包含产品或人工的询价，询价单将由配送中心审核！", vbQuestion + vbYesNo + vbDefaultButton2, "请注意！")
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.workKK
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "合同评审单"
    mod1.cmd.Parameters("@NBLX") = "添加"
    mod1.cmd.Parameters("@bh") = ""
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtXMMC.Text) '项目名称
    mod1.cmd.Parameters("@mt2") = mod1.DName '业务员
    mod1.cmd.Parameters("@mt3") = mod1.DHid '创建日期
    mod1.cmd.Parameters("@mt4") = mod1.Qy
    mod1.cmd.Parameters("@mt5") = mod1.Bm
    mod1.cmd.Parameters("@mt6") = Trim(txtKhmc.Text) '客户名称
    mod1.cmd.Parameters("@mt7") = Trim(txtKhmc.ToolTipText) '客户代号
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = ""
    mod1.cmd.Parameters("@mt11") = ""
    mod1.cmd.Parameters("@mt12") = ""
    mod1.cmd.Parameters("@mt13") = ""
    mod1.cmd.Parameters("@mt14") = ""
    mod1.cmd.Parameters("@mt15") = ""
    mod1.cmd.Parameters("@mt16") = ""
    mod1.cmd.Parameters("@mt17") = ""
    mod1.cmd.Parameters("@mt18") = ""
    mod1.cmd.Parameters("@mt19") = ""
    mod1.cmd.Parameters("@mt20") = ""
    mod1.cmd.Parameters("@mt21") = ""
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtXMMC.BoundText)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 0
    mod1.cmd.Parameters("@mm11") = 0
    mod1.cmd.Parameters("@mm12") = 0
    mod1.cmd.Parameters("@mm13") = 0
    mod1.cmd.Parameters("@mm14") = 0
    mod1.cmd.Parameters("@mm15") = 0
    mod1.cmd.Parameters("@mm16") = 0
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
''''''    If PRF = vbYes Then
''''''        mod1.cmd.Parameters("@mm19") = 1 '纯配件合同还是其它合同
''''''    ElseIf PRF = vbNo Then
        mod1.cmd.Parameters("@mm19") = 2 '纯配件合同还是其它合同
'''''    End If
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = Null
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
   ' MsgBox "b"
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 1 Then
            cmdNew.Enabled = False
        End If
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If

    
Set mod1.cmd = Nothing
'fmxc.Show

txtKhmc.Enabled = False

End Sub

Private Sub cmdRef_Click()
Dim tt As String
On Error Resume Next
If ZT <> "作废" Then
    DT = "Select * from htView where ((业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
    "')) and 状态='" & ZT & "' and 合同日期>='" & dd1.Value & "' and 合同日期<='" & dd2.Value & "' order by 合同日期 desc"
Else
    DT = "Select * from htViewDel where ((业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
    "')) and 合同日期>='" & dd1.Value & "' and 合同日期<='" & dd2.Value & "' order by 合同日期 desc"
End If
'''    htBrow.adoBr.Close
'''    htBrow.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''    Set htBrow.dtgBr.DataSource = htBrow.adoBr
'''    htBrow.dtgBr.FixedRows = 0
'''    htBrow.dtgBr.MergeCol(1) = True
'''    htBrow.dtgBr.MergeCol(2) = True
'''    htBrow.dtgBr.MergeCol(3) = True
'''    htBrow.dtgBr.MergeCol(4) = True
'''    htBrow.dtgBr.MergeCol(7) = True
'''    htBrow.dtgBr.MergeCells = 3
'''    htBrow.dtgBr.FixedRows = 1
    Call dtgREF
End Sub


Private Sub cmdRef1_Click()
Dim tt As String
On Error Resume Next
Select Case comXZ.Text
    Case "合同金额"
        DT = "Select * from htView where ((业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
        "')) and 合同金额=" & Val(txtYc.Text) & " order by 合同日期 desc"
    Case "项目名称"
        DT = "Select * from htView where ((业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
        "')) and 项目名称 like '%" & Trim(txtYc.Text) & "%' order by 合同日期 desc"
    Case "合同编号"
        DT = "Select * from htView where ((业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or ( Xuid='" & mod1.DHid & _
        "')) and 合同编号 like '%" & Trim(txtYc.Text) & "%' order by 合同日期 desc"
End Select

'''    htBrow.adoBr.Close
'''    htBrow.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''    Set htBrow.dtgBr.DataSource = htBrow.adoBr
'''    If htBrow.adoBr.RecordCount > 0 Then
'''        htBrow.dtgBr.FixedRows = 0
'''        htBrow.dtgBr.MergeCol(1) = True
'''        htBrow.dtgBr.MergeCol(2) = True
'''        htBrow.dtgBr.MergeCol(3) = True
'''        htBrow.dtgBr.MergeCol(4) = True
'''        htBrow.dtgBr.MergeCol(7) = True
'''        htBrow.dtgBr.MergeCells = 3
'''        htBrow.dtgBr.FixedRows = 1
'''    End If
    Call dtgREF
End Sub


Private Sub cmdRun_Click()
If optXJ.Value = True And lblLx.Caption = "" Then
    MsgBox "请确定询价单的类型!"
    FmxcLx.Show
    FmxcLx.ZOrder 0
    Exit Sub
End If
If optHT.Value = True Then
    Call FMXCXmmc.Qing
    FMXCXmmc.Show
    FMXCXmmc.ZOrder 0
End If
End Sub

Private Sub cmdWZX_Click()
Dim tt As String
On Error Resume Next
    DT = "Select * from htView3 where (ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (项目归属人='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
    "') order by qrq desc"

    Call dtgREF
End Sub

Private Sub cmdXjd_Click()
'新版本2013
    FmxcLxNew.Show
    FmxcLxNew.cmdNew.Caption = "生成询价单"
    FmxcLxNew.cmdNew.ToolTipText = 0
    FmxcLxNew.LX = ""
    FmxcLxNew.ZOrder 0

''''旧版本2012
'''''    FmxcLx.Show
'''''    FmxcLx.cmdNew.Caption = "生成询价单"
'''''    FmxcLx.cmdNew.ToolTipText = 0
'''''    FmxcLx.LX = ""
'''''    FmxcLx.ZOrder 0
End Sub

Private Sub cmdXJView_Click()
Me.Enabled = True
    mod1.BTZ = 36
    frmGxBiao.Visible = False
    tt = "select * from xunjiaView where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "' order by bid desc"
    On Error Resume Next
    frmGxBiao.adoXj.Close
    frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
    If frmGxBiao.adoXj.RecordCount > 1 Then
        frmGxBiao.dtgXj.FixedRows = 0
        frmGxBiao.dtgXj.MergeCol(1) = True
        frmGxBiao.dtgXj.MergeCol(3) = True
        frmGxBiao.dtgXj.MergeCol(4) = True
        frmGxBiao.dtgXj.MergeCells = 3
        frmGxBiao.dtgXj.FixedRows = 1
    End If
    frmGxBiao.Visible = True
'''''''''    '取得新建维保询价单及购销询价单的流程参数
'''''''''    tt = "xunJiaBut('" & mod1.DName & "','" & mod1.DHid & "','维保询价')"
'''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
'''''''''    frmGxBiao.cmdNew.Tag = mod1.HTP.Fields("nlb").Value
'''''''''    frmGxBiao.cmdNew.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
'''''''''    frmGxBiao.cmdDx.Tag = mod1.HTP.Fields("nlb").Value                          '大修的流程同维保
'''''''''    frmGxBiao.cmdDx.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
'''''''''    tt = "xunJiaBut('" & mod1.DName & "','" & mod1.DHid & "','购销')"
'''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
'''''''''    frmGxBiao.cmdCreat.Tag = mod1.HTP.Fields("nlb").Value
'''''''''    frmGxBiao.cmdCreat.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
'''''''''    frmGxBiao.cmdCP.Tag = mod1.HTP.Fields("nlb").Value
'''''''''    frmGxBiao.cmdCP.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
'''''''''    frmGxBiao.cmdFb.Tag = mod1.HTP.Fields("nlb").Value
'''''''''    frmGxBiao.cmdFb.ToolTipText = "流程总数为:" & mod1.HTP.Fields("lcou").Value
'''''''''    frmGxBiao.frmNew.Visible = True
'''''''''    frmGxBiao.frmC.Visible = False
End Sub

Private Sub Command2_Click()
Dim tt As String
On Error Resume Next
    DT = "Select * from htView where (业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
    "') order by 合同日期 desc"
'''''    htBrow.adoBr.Close
'''''    htBrow.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''    Set htBrow.dtgBr.DataSource = htBrow.adoBr
'''''    If htBrow.adoBr.RecordCount > 0 Then
'''''        htBrow.dtgBr.FixedRows = 0
'''''        htBrow.dtgBr.MergeCol(1) = True
'''''        htBrow.dtgBr.MergeCol(2) = True
'''''        htBrow.dtgBr.MergeCol(3) = True
'''''        htBrow.dtgBr.MergeCol(4) = True
'''''        htBrow.dtgBr.MergeCol(7) = True
'''''        htBrow.dtgBr.MergeCells = 3
'''''        htBrow.dtgBr.FixedRows = 1
'''''    End If
    Call dtgREF
End Sub

''''''Private Sub dtgBr_DblClick()
''''''Static Px As Boolean
''''''
''''''If dtgBr.Row = 1 Then
''''''    If Px = True Then
''''''        dtgBr.Sort dtgBr.Col, SortAscending
''''''        Px = False
''''''    Else
''''''        dtgBr.Sort dtgBr.Col, SortDescending
''''''        Px = True
''''''    End If
'''''''Else
'''''''    MsgBox MGa.ColData(1)
''''''End If
''''''
''''''End Sub


Private Sub dtgBr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static ZF As Boolean
'dtgBr.Col = 2
'txtXMMC.Text = dtgBr.Text
If Button <> 2 Then Exit Sub
If ZF = False Then
        
        Me.dtgBr.FixedRows = 0
        Me.dtgBr.MergeCol(1) = True
        Me.dtgBr.MergeCol(2) = True
        Me.dtgBr.MergeCol(3) = True
        Me.dtgBr.MergeCol(4) = True
        Me.dtgBr.MergeCol(7) = True
        Me.dtgBr.MergeCol(13) = True
        Me.dtgBr.MergeCells = 0
        Me.dtgBr.FixedRows = 1
        ZF = True
Else
        Me.dtgBr.FixedRows = 0
        Me.dtgBr.MergeCol(1) = True
        Me.dtgBr.MergeCol(2) = True
        Me.dtgBr.MergeCol(3) = True
        Me.dtgBr.MergeCol(4) = True
        Me.dtgBr.MergeCol(7) = True
        Me.dtgBr.MergeCol(13) = True
        Me.dtgBr.MergeCells = 3
        Me.dtgBr.FixedRows = 1
        ZF = False
End If

End Sub

Private Sub Form_Load()
timWait.Enabled = False
timQuit.Enabled = False
Me.Left = 0
Me.Top = 0

dd1.Value = #1/1/2003#
dd2.Value = #12/31/2007#
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight

Set adoBr = CreateObject("adodb.recordset")
dtgBr.ColWidth(0) = 300
dtgBr.ColWidth(2) = 3900
dtgBr.ColWidth(4) = 1000
dtgBr.ColWidth(6) = 2100
dtgBr.ColWidth(7) = 1000
dtgBr.ColWidth(8) = 0 'hid
dtgBr.ColWidth(9) = 0 'uid
dtgBr.ColWidth(10) = 0 'Xywy
dtgBr.ColWidth(11) = 0 'Xuid
dtgBr.ColWidth(12) = 0 'NewF
dtgBr.ColWidth(13) = 0 '部门
dtgBr.ColWidth(14) = 0 'comid
dtgBr.ColWidth(15) = 0 'khdh
dtgBr.ColWidth(17) = 0 'fid
'If mod1.MName = "马晓聪" Then
    frmNew.Visible = True
'Else
'    frmNew.Visible = False
'End If
txtKhmc.Enabled = False
Set adoXmmc = CreateObject("adodb.recordset")
Set adoKhmc = CreateObject("adodb.recordset")
PRF = 0

'NF.Left = 11190
'NF.Top = 5700
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
Me.Visible = False
frmZu.Enabled = True
Cancel = True
frmZu.SetFocus
End If
End Sub

Private Sub NiceOption1_Click()
FmxcLx.Show

lblLx.Caption = ""
End Sub

Private Sub NiceButton1_Click()
Dim tt As String
Dim Ra


 timZm = 2
 Set mod1.cmd = CreateObject("adodb.command")
 mod1.cmd.ActiveConnection = mod1.workKK
 mod1.cmd.CommandText = "MLAdd"
 mod1.cmd.CommandType = adCmdStoredProc
 mod1.cmd.Parameters("@zid") = 0
 mod1.cmd.Parameters("@errch") = ""
 mod1.cmd.Parameters("@NB") = "新合同2013"
 mod1.cmd.Parameters("@NBLX") = "添加询价单"
 mod1.cmd.Parameters("@bh") = ""
 mod1.cmd.Parameters("@ywy") = mod1.DName
 mod1.cmd.Parameters("@uid") = mod1.DHid
 mod1.cmd.Parameters("@mt1") = ""
 mod1.cmd.Parameters("@mt2") = "询价指令"
 mod1.cmd.Parameters("@mt5") = ""
 mod1.cmd.Parameters("@mt25") = ""
 mod1.cmd.Parameters("@mlt1") = ""
 mod1.cmd.Parameters("@mm1") = 0
 mod1.cmd.Parameters("@mm2") = 0
' Exit Sub
 mod1.cmd.Parameters("@md1") = Null
 Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
 mod1.cmd.Execute
' MsgBox "b"
 mod1.Zid = mod1.cmd.Parameters("@zid").Value
 If mod1.cmd.Parameters("@errch").Value <> "成功" Then
     MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
     If timZm = 1 Then
         cmdNew.Enabled = False
     End If
     Exit Sub
 Else '提交成功,等待系统中心处理数据
     Me.Enabled = False
     frmWaitA.Visible = True
     frmWaitA.Timer2.Enabled = False

     frmWaitA.ZOrder 0
     frmWaitA.Timer2.Enabled = True
     timWait.Enabled = True
 End If
 Set mod1.cmd = Nothing

End Sub

Private Sub OKButton_Click()
mod1.BTZ = 6
Dim tt As String
Dim xZ As String
Dim NewF As Integer
Dim Hid As Long
'Dim Lid As String
On Error Resume Next
dtgBr.Col = 4
xZ = dtgBr.Text
dtgBr.Col = 8
Hid = dtgBr.Text
dtgBr.Col = 11
If Val(dtgBr.Text) = 0 Then
dtgBr.Col = 12
End If
NewF = dtgBr.Text

'Lid = Str(Lid)
If mod1.DKZ(Hid, 1) = True Then
        MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
        Exit Sub
End If

frmWait.Visible = True
frmWait.ZOrder 0
frmWait.Refresh
'htBrow.MousePointer = 11
Me.Enabled = False
'mod1.MPld = False '初始化,不生成配料单
If NewF = 0 Then
    If xZ = "C. 维保合同" Or xZ = "D. 维修合同" Then
    'mod1.comJZ = False
    wbHTP.Visible = False
    Call modHt.wbQing
    
    
    tt = "Select * from htping where hid=" & Hid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Call modHt.wbBound
    
    
    '打开材料表
    tt = "Select * from htSale where htbh='" & wbHTP.lblHid.Caption & "'"
    wbMx.adoRGF.Recordset.Close
    wbMx.adoRGF.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set wbMx.dtgSale.DataSource = wbMx.adoRGF
    wbMx.lblChg.Caption = wbHTP.txtClcb1.Text
    
    '打开应收款表
    tt = "Select * from htping1 where htBh='" & wbHTP.lblHid.Caption & "' order by rq"
    frmFuK.adoHpt.Recordset.Close
    frmFuK.adoHpt.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set wbMx.dtgFk.DataSource = frmFuK.adoHpt
    
    '打开佣金表
    tt = "Select * from Yongjin where htBh='" & wbHTP.txtHtbh.Text & "' order by yId"
    frmYJ.adoYj.Recordset.Close
    frmYJ.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set frmYJ.dtgYJ.DataSource = frmYJ.adoYj
    
    ''打开出工信息表(如果为评审阶段则不显示）
    'If wbHTP.optZ.Value = True Or wbHTP.optW.Value = True Then
    '    tt = "Select max(gzb.rq),max(gzb.wxWorker),sum(workXX.wTime),max(bhid)" & _
    '    "max(htbh) from gzb cross join workXX where gzb.bhid=workXX.bhid and gzb.htBh='" & _
    '    wbHTP.txtHtbh.Text & "' group by gzb.bhid"
    '    form2Htp.adoGzb.Recordset.Close
    '    form2Htp.adoGzb.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '    Set wbMx.dtgGzb.DataSource = form2Htp.adoGzb
    'End If
    wbHTP.Visible = True
    
    wbHTP.txtYj1.Visible = False
    wbHTP.txtYj2.Visible = False
    wbHTP.txtLr1.Visible = False
    wbHTP.txtLr2.Visible = False
    wbHTP.lblTcBe.Visible = False
    wbHTP.txtTcBe.Visible = False
    wbHTP.UpDa.Visible = False
    wbHTP.lblYj.Visible = False
    wbHTP.lblLR.Visible = False
    wbHTP.lblTC.Visible = False
    Exit Sub
    End If
    
    
    
    
    
    
    
    
    
    
    '购销合同
    
    form2Htp.Visible = True
    mod1.workTt = ""
    mod1.workTt = "Select * from htPing where hid=" & Hid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open mod1.workTt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    form2Htp.lblHtxz.Caption = ""
    
    Call modHt.htQing
    Call modHt.htBound '绑定合同评审单字段
    
    '如果维修合同，则计算总工时，并列出出工列表
    'If form2Htp.optA(1).Value = True Or form2Htp.optA(3).Value = True Or form2Htp.optA(4).Value = True Then
    
        
        

    
    
    '打开收款表
    
    
    tt = "Select * from htPing1 where htBh='" & form2Htp.lblHid.Caption & "' order by rq"
    frmFuK.adoHpt.Recordset.Close
    frmFuK.adoHpt.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    
    
    Set frmFuK.dtgFk.DataSource = frmFuK.adoHpt
    
    'ft = "Select * from yiFk Where htBh='" & frmFuK.adoHpt.Recordset.Fields("htBh").Value & _
    '"' and yingRQ='" & frmFuK.adoHpt.Recordset.Fields("rq").Value & "' order by yiRq"
    'frmFuK.adoYf.Recordset.Close
    'frmFuK.adoYf.Recordset.Open ft, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'Set frmFuK.dtgYf.DataSource = frmFuK.adoYf
    
    '打开产品表
    tt = ""
    tt = "Select * from htSale Where htBh='" & form2Htp.txtHtbh.Text & "'"
    form2Htp.adoSale.Recordset.Close
    form2Htp.adoSale.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set form2Htp.dtgSale.DataSource = form2Htp.adoSale
    Set form2Htp.dtgYJ.DataSource = form2Htp.adoSale
    Set form2Htp.dtgZj.DataSource = form2Htp.adoSale
    
    ''打开“取自库存表”
    'tt = "Select * from kcJa where htBh='" & form2Htp.txtHtbh.Text & "'"
    'form2Htp.adoKu.Recordset.Close
    'form2Htp.adoKu.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'Set form2Htp.dtgKu.DataSource = form2Htp.adoKu
    
    ''打开采购表
    'ft = "Select * from CG Where htbh='" & form2Htp.txtHtbh.Text & "' and khmc<>'库存'"
    'frmAdo.adoTmp.Recordset.Close
    'frmAdo.adoTmp.Recordset.Open ft, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'Set form2Htp.dtgCG.DataSource = frmAdo.adoTmp
    
    '打开佣金表
    tt = "Select * from Yongjin where htBh='" & form2Htp.txtHtbh.Text & "' order by yId"
    frmYJ.adoYj.Recordset.Close
    frmYJ.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set frmYJ.dtgYJ.DataSource = frmYJ.adoYj
    
    
    
    
    form2Htp.tabHt.TabEnabled(1) = True
    form2Htp.tabHt.TabEnabled(2) = True
    'End If
    
    
    
    
    
    
    
    form2Htp.tabHt.Tab = 0
    Me.MousePointer = 0
    
    
        '佣金、利润2、提成不显示
        form2Htp.txtYj1.Visible = False
        form2Htp.txtYj2.Visible = False
        form2Htp.txtLr1.Visible = False
        form2Htp.txtLr2.Visible = False
        'form2Htp.txtTc1.Visible = False
        'form2Htp.txtTc2.Visible = False
        form2Htp.lblYj.Visible = False
        form2Htp.lblLr2.Visible = False
        'form2Htp.lblTc.Visible = False
ElseIf NewF = 1 Then
        Call modHt.NewQing
        
        Call modHt.NewBound(Hid)

        frmWbNew.Visible = True
ElseIf NewF = 2 Then
        Call modNewHT.NewMQing
        
        Call modNewHT.NewMBound(Hid)
        FMXC.lblMQM(0).Visible = True
        FMXC.lblMTm(0).Visible = True
        FMXC.cmdMQm(0).Visible = True
ElseIf NewF = 3 Or NewF = 5 Or NewF = 7 Then
        Call modNewHT.NewMQing
        
        Call modNewHT.NewB(Hid)
        FMXC.lblMQM(0).Visible = True
        FMXC.lblMTm(0).Visible = True
        FMXC.cmdMQm(0).Visible = True
ElseIf NewF = 6 Or NewF = 8 Then
    Call FmxcNew.Bound(Hid)
    FmxcNew.Show
    FmxcNew.ZOrder 0

End If
FmxcNew.Width = mod1.FWidth + 500
FmxcNew.Height = mod1.FHeight
FmxcNew.frmNewLx.Left = 5070
FmxcNew.frmNewLx.Top = 0
End Sub

Private Sub optXJ_Click()
    FmxcLx.Show
    FmxcLx.ZOrder 0
End Sub

Private Sub optY_Click(Index As Integer)
ZT = optY(Index).Caption
End Sub





Private Sub Timer1_Timer()

End Sub


Private Sub timQuit_Timer()
On Error Resume Next
Dim ii As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '如果为添加合同评审
FMXC.Show
FMXC.txtXMMC.Text = txtXMMC.Text
FMXC.txtXMMC.ToolTipText = txtXMMC.BoundText
FMXC.txtKhmc.Text = txtKhmc.Text
FMXC.txtKhdm.Text = txtKhmc.ToolTipText
FMXC.cmdWb.ToolTipText = txtXMMC.BoundText
FMXC.comQy.Text = mod1.Qy
FMXC.comKQY.Text = mod1.Qy
FMXC.txtYwy.Text = mod1.DName
FMXC.txtYwy.ToolTipText = mod1.DHid
FMXC.txtXYwy.Text = mod1.DName
FMXC.txtXYwy.ToolTipText = mod1.DHid
FMXC.txtHtrq.Text = mod1.DQda
FMXC.txtHtbh.Text = "HMNEW"
FMXC.lblMHid.Caption = Hid
FMXC.frmYJ.Visible = False
FMXC.lblLc.Caption = 0
FMXC.lblLcRen.Caption = mod1.DName
FMXC.lblLcUid.Caption = mod1.DHid
FMXC.cmdSave.Enabled = True
FMXC.cmdHt.Visible = True
FMXC.frmFk.Visible = True
FMXC.dtgSD.Visible = False
'打开应收款表
tt = "select * from htFk where htbh='" & FMXC.lblMHid.Caption & "'"
mod1.mFk.Close
mod1.mFk.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''If IsNull(fmxc.adoFfk.RecordCount) = True Then
''    MsgBox ("读取数据错误!")
''    Exit Sub
''End If
Set FMXC.MMdtgFk.DataSource = mod1.mFk

FMXC.txtW5.Locked = False
FMXC.txtW6.Locked = False
'''If mod1.Qy = "上海" Then
    FMXC.txtW3.Locked = True
    'FMXC.cmdW3.Visible = True
    FMXC.txtW4.Locked = True
    'FMXC.cmdW4.Visible = True
'''Else
'''    FMXC.txtW3.Locked = False
'''    FMXC.cmdW3.Visible = False
'''    FMXC.txtW4.Locked = False
'''    FMXC.cmdW4.Visible = False
'''End If

'版本切换
FMXC.lblWC.Visible = False
FMXC.txtQt1.Visible = False
FMXC.txtW5.Visible = False
FMXC.txtH5.Left = 1890
FMXC.txtH5.Width = 2175
FMXC.txtW6.Visible = False
FMXC.txtH6.Left = 1890
FMXC.txtH6.Width = 2175
FMXC.lblYug.Caption = "基准价"
FMXC.lblYug2.Visible = False
FMXC.chkA.ForeColor = &HC00000
FMXC.chkB.ForeColor = &HC00000
FMXC.chkC.ForeColor = &HC00000
FMXC.chkD.ForeColor = &HC00000
FMXC.chkE.ForeColor = &HC00000
FMXC.chkF.ForeColor = &HC00000
FMXC.txtH1.ForeColor = &HC00000
FMXC.txtH2.ForeColor = &HC00000
FMXC.txtW3.ForeColor = &HC00000
FMXC.txtW4.ForeColor = &HC00000
FMXC.txtH5.ForeColor = &HC00000
FMXC.txtH6.ForeColor = &HC00000
FMXC.lblWC.Visible = False: FMXC.txtQt1.Visible = False
FMXC.lblCBZE.Caption = "基准总价": FMXC.lblCBZE.ForeColor = &HC00000
FMXC.txtCbze1.Width = 2475: FMXC.txtCbze2.Visible = False
FMXC.txtClcb1.Width = 2475: FMXC.txtClcb2.Visible = False
FMXC.txtFbje1.Width = 2475: FMXC.txtFbje2.Visible = False
FMXC.lblCL.Visible = False: FMXC.txtCLF1.Visible = False
FMXC.lblCb.Visible = False
FMXC.lblYug.ForeColor = FMXC.chkA.ForeColor
FMXC.lblClcb.Top = 1650: FMXC.txtClcb1.Top = 1650: FMXC.txtClcb2.Top = 1650
FMXC.lblRG.Top = 2200: FMXC.txtRgf1.Top = 2200: FMXC.txtRGF2.Top = 2200
ElseIf timZm = 2 Then
    Call FmxcXJ.Bound(Bid)
    FmxcXJ.Show
    FmxcXJ.ZOrder 0
    MsgBox "请在备注中填写所要询价的内容！"
End If
timQuit.Enabled = False
Hid = 0
Me.Enabled = True
''''''''If PRF = vbYes Then
''''''''    FMXC.cmdW1.Visible = False: FMXC.cmdW2.Visible = False: FMXC.cmdW3.Visible = False: FMXC.cmdW4.Visible = False: FMXC.cmdW6.Visible = False
''''''''ElseIf PRF = vbNo Then
''''''''    FMXC.cmdW1.Visible = True: FMXC.cmdW2.Visible = True: FMXC.cmdW3.Visible = True: FMXC.cmdW4.Visible = True: FMXC.cmdW6.Visible = True
''''''''End If
FMXC.dtgFL.Col = 4
For oo = 1 To 7
    FMXC.dtgFL.Row = oo
        If oo < 6 Then
            FMXC.dtgFL.Text = "双击新增"
        ElseIf oo = 6 Then
                FMXC.dtgFL.Text = "双击新增  "
        ElseIf oo = 7 Then
                FMXC.dtgFL.Text = "双击新增     "
        End If
Next
    FMXC.dtgFL.Visible = True
    FMXC.dtgFL.MergeCol(4) = True
    FMXC.dtgFL.MergeCells = flexMergeRestrictColumns
'''''If mod1.DName <> "谢雪梅" Then
'''''    FMXC.dtgFL.Visible = False
'''''End If
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    If timZm = 2 Then
        Bid = Val(mod1.WP.Fields("bh").Value)
    End If
    mod1.Ti = 5
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    timWait.Enabled = False
    If timZm = 1 Then
        Hid = mod1.WP.Fields("bh").Value
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    If timZm = 1 Then
        cmdNew.Enabled = False
    End If
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdNew.Enabled = False
    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub

Private Sub txtKhmc_Click()
Dim tt As String
On Error Resume Next
If Me.Visible = False Then Exit Sub

tt = "Select khdh from khzl where khqc ='" & txtKhmc.Text & "'  order by kid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
txtKhmc.ToolTipText = mod1.HTP.Fields("khdh").Value

End Sub


Private Sub txtXMMC_Click(Area As Integer)
Dim tt As String
Dim oo As Integer

On Error Resume Next
If Me.Visible = False Then Exit Sub


    'tt = "select khqc,khdh from khzl where xid=" & txtXMMC.BoundText
    tt = "Select yzmc,wymc,qt1mc,qt2mc,qt3mc,qt4mc,qt5mc from xmKhmc where xid=" & Val(txtXMMC.BoundText)

    adoKhmc.Close
    adoKhmc.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    Set txtKhmc.RowSource = adoKhmc
'    txtKhmc.ListField = "khqc"
'    txtKhmc.BoundColumn = "khdh"
For oo = 6 To 0 Step -1
    txtKhmc.RemoveItem oo
    'KhId(oo) = 0
Next
    
    

    If adoKhmc.RecordCount = 1 Then
        If IsNull(adoKhmc.Fields("yzmc").Value) = False Then
            txtKhmc.AddItem adoKhmc.Fields("yzmc").Value
            'KhId(0) = adoKhmc.Fields("yzid").Value
        End If
        If IsNull(adoKhmc.Fields("wymc")) = False Then
            txtKhmc.AddItem adoKhmc.Fields("wymc").Value
            'KhId(1) = adoKhmc.Fields("wyid").Value
        End If
        For oo = 1 To 5
            If IsNull(adoKhmc.Fields("qt" & oo & "mc")) = False And adoKhmc.Fields("qt" & oo & "mc") <> "" Then
                txtKhmc.AddItem adoKhmc.Fields("qt" & oo & "mc").Value
                'KhId(oo + 1) = adoKhmc.Fields("qt" & oo & "id").Value
            End If
        Next
    End If
    adoKhmc.Close
    If txtKhmc.ListCount > 0 Then
        txtKhmc.Enabled = True
        lblTX.Visible = False
    Else
        txtKhmc.Enabled = False
        lblTX.Visible = True
        'MsgBox ("您的项目资料有问题，可能没有建立客户档案，请查实！")
    End If
End Sub

Private Sub txtXmmc_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tt As String

On Error Resume Next
If Me.Visible = False Then Exit Sub

If KeyCode = 13 And txtXMMC.Text <> "" Then
    tt = "select xmmc,xid from xmzl where uid='" & mod1.DHid & "' and xmmc like '%" & txtXMMC.Text & "%'"
    adoXmmc.Close
    adoXmmc.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set txtXMMC.RowSource = adoXmmc
    txtXMMC.ListField = "xmmc"
    txtXMMC.BoundColumn = "xid"
    
    
End If
End Sub


Private Sub txtYc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmdRef1_Click
End If
End Sub



Public Sub dtgGG()
dtgBr.Clear
dtgBr.Row = 0
dtgBr.Col = 1: dtgBr.Text = "业务员"
dtgBr.Col = 2: dtgBr.Text = "项目名称"
dtgBr.Col = 3: dtgBr.Text = "合同日期"
dtgBr.Col = 4: dtgBr.Text = "合同性质"
dtgBr.Col = 5: dtgBr.Text = "合同金额"
dtgBr.Col = 6: dtgBr.Text = "合同编号"
dtgBr.Col = 7: dtgBr.Text = "状态"
dtgBr.Col = 16: dtgBr.Text = "项目管理者"
End Sub

Public Sub dtgREF()
Dim oo As Integer
Dim ii As Integer
Dim YY As Integer
On Error Resume Next
    dtgBr.Visible = False
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open Me.DT, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2)
   Call htBrow.dtgGG
    htBrow.dtgBr.Rows = La + 30
    htBrow.dtgN.Rows = La + 30: htBrow.dtgN.Cols = htBrow.dtgBr.Cols
    For oo = 1 To La + 1
        htBrow.dtgBr.Row = oo: htBrow.dtgN.Row = oo
        For ii = 1 To htBrow.dtgBr.Cols
            htBrow.dtgBr.Col = ii: htBrow.dtgN.Col = ii
            If ii = 3 Then '日期格式化
                htBrow.dtgBr.Text = Format(Ra(ii - 1, oo - 1), "YYYY-MM-DD")
                htBrow.dtgN.Text = Format(Ra(ii - 1, oo - 1), "YYYY-MM-DD")
            Else
                htBrow.dtgBr.Text = Ra(ii - 1, oo - 1)
                htBrow.dtgN.Text = Ra(ii - 1, oo - 1)
            End If
            If ii = 17 And Val(htBrow.dtgBr.Text) > 0 Then
                For YY = 1 To htBrow.dtgBr.Cols
                    htBrow.dtgBr.Col = YY
                    htBrow.dtgBr.CellForeColor = &H8000000D
                Next
            End If
            htBrow.dtgBr.Col = ii
        Next
    Next
    dtgBr.Visible = True
End Sub
