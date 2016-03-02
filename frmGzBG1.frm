VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGzBG 
   Caption         =   "项目实施状况"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   FillStyle       =   5  'Downward Diagonal
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdNew 
      Caption         =   "新工作报告"
      Height          =   495
      Left            =   13470
      TabIndex        =   40
      Top             =   8640
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7710
      Top             =   8490
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8520
      Top             =   8490
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgKh 
      Height          =   7575
      Left            =   0
      TabIndex        =   35
      Top             =   690
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   13361
      _Version        =   393216
      BackColor       =   12648447
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5745
      Left            =   4890
      TabIndex        =   25
      Top             =   30
      Width           =   10305
      Begin VB.CommandButton cmdLz 
         Caption         =   "下一周"
         Height          =   285
         Left            =   9450
         TabIndex        =   29
         Top             =   5460
         Width           =   795
      End
      Begin VB.CommandButton cmdPz 
         Caption         =   "上一周"
         Height          =   285
         Left            =   8730
         TabIndex        =   28
         Top             =   5460
         Width           =   705
      End
      Begin MSDataGridLib.DataGrid dtgJi 
         Bindings        =   "frmGzBG1.frx":0000
         Height          =   4695
         Left            =   270
         TabIndex        =   26
         Top             =   660
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   8281
         _Version        =   393216
         ForeColor       =   -2147483647
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "工作计划"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "aTime"
            Caption         =   "日期"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "YYYY-M-D aaaa"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "xmmc"
            Caption         =   "计划执行项目"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "NewF"
            Caption         =   "NewF"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Gid"
            Caption         =   "Gid"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpXr 
         Height          =   315
         Left            =   7350
         TabIndex        =   27
         Top             =   5430
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   16711680
         Format          =   108986369
         CurrentDate     =   38272
      End
      Begin MSDataGridLib.DataGrid dtgXmgz 
         Bindings        =   "frmGzBG1.frx":0014
         Height          =   4695
         Left            =   4680
         TabIndex        =   30
         Top             =   660
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   8281
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "您的销售经历"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "aTime"
            Caption         =   "日期"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "YYYY-M-D aaaa"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "khQc"
            Caption         =   "拜访客户"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "xmFy"
            Caption         =   "项目费用"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "NewF"
            Caption         =   "NewF"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Gid"
            Caption         =   "Gid"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2805.166
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label Label10 
         Caption         =   "一周实施情况"
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
         Left            =   2640
         TabIndex        =   36
         Top             =   120
         Width           =   2205
      End
      Begin VB.Label lblFr 
         Height          =   255
         Left            =   4410
         TabIndex        =   34
         Top             =   5460
         Width           =   885
      End
      Begin VB.Label Label4 
         Caption         =   "~~"
         Height          =   165
         Left            =   5310
         TabIndex        =   33
         Top             =   5520
         Width           =   225
      End
      Begin VB.Label lblLr 
         Height          =   255
         Left            =   5580
         TabIndex        =   32
         Top             =   5460
         Width           =   885
      End
      Begin VB.Label Label5 
         Caption         =   "选定日期"
         Height          =   195
         Left            =   6600
         TabIndex        =   31
         Top             =   5490
         Width           =   735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMx 
      Height          =   2385
      Left            =   5220
      TabIndex        =   23
      Top             =   5910
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   4207
      _Version        =   393216
      BackColor       =   16777088
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查询"
      Height          =   645
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8460
      Width           =   1065
   End
   Begin MSComCtl2.DTPicker dtpR 
      Height          =   255
      Left            =   630
      TabIndex        =   19
      Top             =   8850
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   108986369
      CurrentDate     =   39036
   End
   Begin MSComCtl2.DTPicker dtpL 
      Height          =   285
      Left            =   630
      TabIndex        =   18
      Top             =   8490
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   108986369
      CurrentDate     =   39036
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "后一月"
      Height          =   285
      Left            =   2310
      TabIndex        =   17
      Top             =   8820
      Width           =   975
   End
   Begin VB.CommandButton cmdPre 
      Caption         =   "前一月"
      Height          =   285
      Left            =   2310
      TabIndex        =   16
      Top             =   8490
      Width           =   975
   End
   Begin VB.CommandButton cmdFw 
      BackColor       =   &H000000FF&
      Caption         =   "选择人员"
      Height          =   315
      Left            =   5310
      TabIndex        =   13
      Top             =   8760
      Width           =   1095
   End
   Begin VB.TextBox comYwy 
      Height          =   345
      Left            =   10710
      TabIndex        =   12
      Text            =   "comYwy"
      Top             =   8700
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "返回"
      Height          =   555
      Left            =   14550
      Picture         =   "frmGzBG1.frx":002A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8610
      Width           =   645
   End
   Begin VB.Frame frmMc 
      BorderStyle     =   0  'None
      Height          =   2325
      Left            =   10530
      TabIndex        =   0
      Top             =   5880
      Width           =   4725
      Begin MSDataGridLib.DataGrid dtgHt 
         Height          =   1005
         Left            =   30
         TabIndex        =   38
         Top             =   870
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   1773
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "合同编号"
            Caption         =   "对应执行合同"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "合同金额"
            Caption         =   "金额"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            RecordSelectors =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   900.284
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "进入"
         Height          =   285
         Left            =   3690
         TabIndex        =   5
         Top             =   1980
         Width           =   1005
      End
      Begin VB.OptionButton optBg 
         Caption         =   "销售日记"
         Height          =   315
         Left            =   3690
         TabIndex        =   9
         Top             =   1560
         Width           =   1065
      End
      Begin VB.OptionButton optJh 
         Caption         =   "计划"
         Height          =   255
         Left            =   3690
         TabIndex        =   8
         Top             =   1200
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtpRq 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "m/d/yy aaaa"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   4
         Top             =   420
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   12582912
         Format          =   108986368
         CurrentDate     =   38272
      End
      Begin VB.ComboBox comKhmc 
         DataField       =   "UserId"
         Height          =   300
         Left            =   1260
         TabIndex        =   2
         ToolTipText     =   "您只能选择客户资料中存在的客户，若为新客户，则必须先在客户资料中添加"
         Top             =   60
         Width           =   3225
      End
      Begin VB.Label lblHtbh 
         Caption         =   "lblHtbh"
         Height          =   225
         Left            =   120
         TabIndex        =   39
         Top             =   2070
         Width           =   2745
      End
      Begin VB.Label lblWeek 
         Caption         =   "五"
         Height          =   225
         Left            =   4260
         TabIndex        =   7
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label3 
         Caption         =   "星期"
         Height          =   225
         Left            =   3870
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "行程日期"
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "项目名称"
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Top             =   90
         Width           =   735
      End
   End
   Begin MSAdodcLib.Adodc adoKhmc 
      Height          =   345
      Left            =   9000
      Top             =   8340
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\work.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\work.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "worker"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label11 
      Caption         =   "项目汇总分析"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      TabIndex        =   37
      Top             =   150
      Width           =   2055
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   3
      X1              =   15240
      X2              =   10440
      Y1              =   8250
      Y2              =   8250
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   3
      X1              =   10470
      X2              =   10470
      Y1              =   8250
      Y2              =   5850
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   3
      X1              =   10440
      X2              =   4830
      Y1              =   5850
      Y2              =   5850
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   3
      X1              =   4830
      X2              =   4830
      Y1              =   0
      Y2              =   5880
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   465
      Left            =   5070
      TabIndex        =   24
      Top             =   7050
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      DrawMode        =   9  'Not Mask Pen
      X1              =   4710
      X2              =   5130
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label Label8 
      Caption         =   "截至"
      Height          =   225
      Left            =   90
      TabIndex        =   21
      Top             =   8910
      Width           =   465
   End
   Begin VB.Label Label7 
      Caption         =   "开始"
      Height          =   225
      Left            =   90
      TabIndex        =   20
      Top             =   8550
      Width           =   435
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      DataField       =   "UserId"
      DataSource      =   "adoKhmc"
      Height          =   345
      Left            =   10410
      TabIndex        =   15
      Top             =   8370
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label lblFw 
      Height          =   285
      Left            =   6480
      TabIndex        =   14
      Top             =   8790
      Width           =   1155
   End
   Begin VB.Label lblYwy 
      Caption         =   "Label6"
      Height          =   315
      Left            =   12450
      TabIndex        =   10
      Top             =   8370
      Visible         =   0   'False
      Width           =   1425
   End
End
Attribute VB_Name = "frmGzBG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AdoKh As Object
Public adoXm As Object
Public adoJi As Object
Public adoMx As Object
Dim adoRq As Object
Dim adoHT As Object
Dim timZm As Integer '数据提交后,由timWait执行的后续命令ID(1添加工作报告)
Private Sub cmdAdd_Click()


End Sub



Private Sub cmdBack_Click()
frmGzBG.Visible = False
frmZu.Enabled = True
End Sub

Private Sub cmdC_Click()
Dim oo As Integer
Dim tt As String
Dim FHg As Single
On Error Resume Next
 tt = "select xmmc as 项目名称,max(khjb) as 平台,sum(xmfy) as 费用 from xmgzB where ywy like '%" & lblFw.Caption & "%' and atime>='" & _
 dtpL.Value & "' and atime<'" & DateSerial(Year(dtpR.Value), Month(dtpR.Value), Day(dtpR.Value) + 1) & "'  group by xmmc order by 项目名称"
Set frmGzBG.AdoKh = CreateObject("adodb.recordset")

frmGzBG.dtgKh.Rows = 2
frmGzBG.AdoKh.Close
frmGzBG.AdoKh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If frmGzBG.AdoKh.RecordCount > 0 Then
    Set frmGzBG.dtgKh.DataSource = frmGzBG.AdoKh
Else
    Set frmGzBG.dtgKh.DataSource = Nothing
End If

'计算总计
FHg = 0
dtgKh.Rows = dtgKh.Rows + 2
dtgKh.Col = 3
For oo = 1 To dtgKh.Rows - 2
    dtgKh.Row = oo
    FHg = FHg + dtgKh.Text
Next
dtgKh.Row = dtgKh.Rows - 1
dtgKh.Text = FHg
dtgKh.Col = 2
dtgKh.Text = "合计"
End Sub

Private Sub cmdFw_Click()
Set Ren.XForm = New frmGzBG
Call mod1.RenXz("frmGzBG", Me, 0)

End Sub

Private Sub cmdGB_Click()

End Sub

Private Sub cmdJh_Click()

End Sub

Private Sub cmdLz_Click()
Dim tt As String
Dim LLR As String
On Error Resume Next
modXmGz.FR = DateSerial(Year(modXmGz.FR), Month(modXmGz.FR), Day(modXmGz.FR) + 7)
modXmGz.LR = DateSerial(Year(modXmGz.LR), Month(modXmGz.LR), Day(modXmGz.LR) + 7)
LLR = DateSerial(Year(modXmGz.LR), Month(modXmGz.LR), Day(modXmGz.LR) + 1)
frmGzBG.lblFR.Caption = modXmGz.FR
frmGzBG.lblLR.Caption = modXmGz.LR
tt = "Select atime,khqc,xmfy,NewF,gid from xmgz where ywy like '%" & lblYwy.Caption & "%' and aTime>='" & modXmGz.FR & _
"' and aTime <'" & LLR & "' and lb=1 order by aTime"
frmGzBG.adoXm.Close
frmGzBG.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''让表格中的日期显示不重复
'Dim Tdate As Date
'frmGzBG.adoXmrq.Recordset.MoveFirst
'Tdate = frmGzBG.adoXmrq.Recordset.Fields("aTime").Value
'Do While Not frmGzBG.adoXmrq.Recordset.EOF
'    frmGzBG.adoXmrq.Recordset.MoveNext
'    If frmGzBG.adoXmrq.Recordset.Fields("aTime").Value = Tdate Then
'        frmGzBG.adoXmrq.Recordset.Fields("aTime").Value = ""
'    Else
'        Tdate = frmGzBG.adoXmrq.Recordset.Fields("aTime").Value
'    End If
'Loop
Set frmGzBG.dtgXmgz.DataSource = frmGzBG.adoXm

tt = "Select atime,xmmc,newF,gid from xmgz where ywy like '%" & lblYwy.Caption & "%' and aTime>='" & modXmGz.FR & _
"' and aTime <='" & modXmGz.LR & "' and lb=0 order by aTime"
frmGzBG.adoJi.Close
frmGzBG.adoJi.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGzBG.dtgJi.DataSource = frmGzBG.adoJi
dtPRQ.Value = lblFR.Caption
lblWeek.Caption = modXmGz.dayWeek(dtPRQ.DayOfWeek)
End Sub

Private Sub cmdNew_Click()

frmGzbN.Show
frmGzbN.dtgB.Row = 1: frmGzbN.dtgB.Col = 0
If frmGzbN.dtgB.Text = "" Then
    frmGzbN.dtgB.Visible = False
    frmGzbN.lblRen.ToolTipText = mod1.DHid   '业务员打开工作报告,默认为本人
    frmGzbN.lblRen.Caption = mod1.DName
    Call frmGzbN.WeekDate(mod1.DQda, mod1.DHid)
  
    Call frmGzbN.QV(frmGzbN.FS)

    frmGzbN.dtgB.Visible = True
    frmGzbN.dtgB.Row = 1
    frmGzbN.cmdXZ.Visible = False

End If
End Sub

Private Sub cmdNext_Click()
Dim Ye As Integer
Ye = Year(dtpL.Value)
dtpL.Value = DateSerial(Year(dtpL.Value), Month(dtpL.Value) + 1, 1)
If Ye = Year(dtpL.Value) Then
    dtpR.Value = DateSerial(Year(dtpR.Value), Month(dtpL.Value) + 1, Day(dtpL.Value) - 1)
Else
    dtpR.Value = DateSerial(Year(dtpR.Value) + 1, Month(dtpL.Value) + 1, Day(dtpL.Value) - 1)
End If
End Sub

Private Sub cmdOk_Click()
Dim tt As String
On Error Resume Next
If optJh.Value = True Then
            Call modXmGz.jhQing
            frmGzJ.Show
            frmGzJ.lblDm.Caption = adoKhmc.Recordset.Fields("khDh").Value
            'frmGzJ.lblXmmc.Caption = adoKhmc.Recordset.Fields("khQc").Value
            frmGzJ.lblRq.Caption = dtPRQ.Value
            frmGzJ.lblWeek.Caption = modXmGz.dayWeek(dtPRQ.DayOfWeek)
            frmGzJ.lblAdr.Caption = adoKhmc.Recordset.Fields("xmAdr").Value
            frmGzJ.lblYwy.Caption = mod1.DName
            
            '获得默认的项目描述、竞争对手、拜访目的、客户平台
            tt = "Select * from xmGz where ywy='" & mod1.DName & "' and khQc='" & _
            adoKhmc.Recordset.Fields("khQc").Value & "' order by Gid"
            frmGzJ.adoXmgz.Recordset.Close
            frmGzJ.adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
            frmGzJ.adoXmgz.Recordset.MoveLast
            frmGzJ.txtXm.Text = frmGzJ.adoXmgz.Recordset.Fields("Xm").Value '项目描述
            frmGzJ.txtjzDC.Text = frmGzJ.adoXmgz.Recordset.Fields("jzDc").Value '竞争对手
            frmGzJ.txtBfMd.Text = frmGzJ.adoXmgz.Recordset.Fields("BfMd").Value '拜访目的
                '客户平台
            If frmGzJ.adoXmgz.Recordset.Fields("khJb").Value = 0 Then
                frmGzJ.optA.Value = True
            ElseIf frmGzJ.adoXmgz.Recordset.Fields("khJb").Value = 30 Then
                frmGzJ.optB.Value = True
            ElseIf frmGzJ.adoXmgz.Recordset.Fields("khJb").Value = 60 Then
                frmGzJ.optC.Value = True
            ElseIf frmGzJ.adoXmgz.Recordset.Fields("khJb").Value = 90 Then
                frmGzJ.optD.Value = True
            End If
            
            '添加空记录，获得Gid
            ''''''''''''''''''''''''''''''''''''''''''''tt = "Select * from xmGz"
            tt = "Select * from xmGz where gid=0"
            frmGzJ.adoXmgz.Recordset.Close
            frmGzJ.adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
            frmGzJ.adoXmgz.Recordset.AddNew "ywy", mod1.DName
            frmGzJ.adoXmgz.Recordset.Update "xmmc", comKhmc.Text
            frmGzJ.adoXmgz.Recordset.UpdateBatch
            modXmGz.Gid = frmGzJ.adoXmgz.Recordset.Fields("gid").Value
            frmGzJ.adoXmgz.Recordset.Close
            frmGzJ.lblKhmc.Caption = comKhmc.Text
            If mod1.KhK = 1 Or mod1.KhK = 2 Then
                frmGzJ.txtzgPd.Locked = False
            Else
                frmGzJ.txtzgPd.Locked = True
            End If
            frmGzJ.frmMod.Enabled = True
            frmGzJ.cmdSave.Enabled = True
            frmGzJ.cmdDel.Enabled = False
            
            
ElseIf optBg.Value = True Then
    If adoHT.RecordCount > 0 And lblHtbh.Caption = "" Then
        MsgBox "请指定费用对应的合同编号!"
        Exit Sub
    End If
    Call modXmGz.xmQing
    frmGzNr.lblHtbh.Caption = lblHtbh.Caption
    tt = "Select * from xmzl where ywy='" & mod1.DName & "' and xmmc='" & comKhmc.Text & "'"
    adoKhmc.Recordset.Close
    adoKhmc.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If adoKhmc.Recordset.RecordCount = 0 Then
        MsgBox "您选择了不正确的客户，请确认！"
    Else
        modXmGz.Ti = True


             
            frmGzNr.Visible = False
'            If comKhmc.Text <> "豪曼" Then
                frmGzNr.lblXid.Caption = adoKhmc.Recordset.Fields("xid").Value '项目代码
                frmGzNr.lblXmmc.Caption = adoKhmc.Recordset.Fields("xmmc").Value '客户名称
                frmGzNr.lblRq.Caption = dtPRQ.Value
                frmGzNr.lblWeek.Caption = modXmGz.dayWeek(dtPRQ.DayOfWeek)
                frmGzNr.lblYwy.Caption = mod1.DName
                frmGzNr.lblCxmFy.Caption = adoKhmc.Recordset.Fields("xmFy").Value
                '客户平台
            If adoKhmc.Recordset.Fields("khJb").Value = 0 Then
                frmGzNr.optA.Value = True
                frmGzNr.optC.Enabled = False
                frmGzNr.optD.Enabled = False
            ElseIf adoKhmc.Recordset.Fields("khJb").Value = 30 Then
                frmGzNr.optC.Enabled = False
                frmGzNr.optD.Enabled = False
                frmGzNr.optB.Value = True
                Call modXmGz.XMPwf
            ElseIf adoKhmc.Recordset.Fields("khJb").Value = 60 Then
                frmGzNr.optC.Value = True
                frmGzNr.optC.Enabled = True
            ElseIf adoKhmc.Recordset.Fields("khJb").Value = 90 Then
                frmGzNr.optD.Value = True
                
            End If
            
        '获得默认的项目描述、竞争对手、拜访目的、客户平台
            tt = "Select xm,jzdc,bfmd from xmGz where xid=" & Val(frmGzNr.lblXid.Caption) & " order by Gid"
            frmGzNr.adoXmgz.Recordset.Close
            frmGzNr.adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
            frmGzNr.adoXmgz.Recordset.MoveLast
            frmGzNr.txtXm.Text = frmGzNr.adoXmgz.Recordset.Fields("Xm").Value '项目描述
            frmGzNr.txtjzDC.Text = frmGzNr.adoXmgz.Recordset.Fields("jzDc").Value '竞争对手
            frmGzNr.txtBfMd.Text = frmGzNr.adoXmgz.Recordset.Fields("BfMd").Value '拜访目的
        
            
        
        '添加空记录，获得Gid
'            tt = "Select * from xmGz order by gid desc"
'            frmGzNr.adoXmgz.Recordset.Close
'            frmGzNr.adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'            frmGzNr.adoXmgz.Recordset.AddNew "ywy", mod1.DName
'            frmGzNr.adoXmgz.Recordset.UpdateBatch
'            'frmGzNr.adoXmgz.Recordset.MoveFirst
'            frmGzNr.lblGid.Caption = frmGzNr.adoXmgz.Recordset.Fields("gid").Value
'            modXmGz.Gid = frmGzNr.adoXmgz.Recordset.Fields("gid").Value
'            frmGzNr.adoXmgz.Recordset.Close
            Set mod1.cmd = CreateObject("adodb.command")
            mod1.cmd.ActiveConnection = mod1.cc
            mod1.cmd.CommandText = "xmGzAdd"
            mod1.cmd.CommandType = adCmdStoredProc
            mod1.cmd.Parameters("@ywy") = mod1.DName
            mod1.cmd.Parameters("@Uid") = mod1.DHid
            mod1.cmd.Parameters("@xmmc") = comKhmc.Text
            mod1.cmd.Parameters("@xid") = Val(frmGzNr.lblXid.Caption)
            mod1.cmd.Parameters("@nLb") = cmdOK.Tag
            mod1.cmd.Parameters("@lcou") = Right(cmdOK.ToolTipText, 1)
            mod1.cmd.Parameters("@lc") = 0
            mod1.cmd.Parameters("@htbh").Value = lblHtbh.Caption
            mod1.cmd.Parameters("@lcRen") = mod1.DName
            mod1.cmd.Parameters("@lcUid") = mod1.DHid
            mod1.cmd.Parameters("@gid").Value = 0
            mod1.cmd.Execute
            'frmGzNr.lblHtbh.Caption = mod1.CMD.Parameters("@htbh").Value
            If mod1.cmd.Parameters("@gid").Value = 0 Then
                MsgBox "网络出现故障,请再试一次,如果还是提交不成功,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
                frmGzNr.Visible = False
                Me.Enabled = True
                Me.ZOrder 0
                Exit Sub
            End If
            frmGzNr.lblGid.Caption = mod1.cmd.Parameters("@gid").Value
            
            Set cmd = Nothing
            frmGzNr.lblLc.Caption = 0
            frmGzNr.lblLcRen.Caption = mod1.DName
            frmGzNr.lblLcUid.Caption = mod1.DHid
            frmGzNr.lblNlb.Caption = frmGxBiao.cmdNew.Tag
            frmGzNr.lblYwy.Caption = mod1.DName
            frmGzNr.lblUid.Caption = mod1.DHid
            frmGzNr.lblLcou.Caption = Right(cmdOK.ToolTipText, 1)
            Set frmGzNr.dtgRen.DataSource = Nothing
            frmGzNr.dtgRen.Refresh
            
        '更新客户交往表和内容
            tt = "select ren,llid from xmren where gid=" & Val(frmGzNr.lblGid.Caption) & " order by llid desc"
            frmGzNr.adoBlx.Close
            frmGzNr.adoBlx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            '    tt = "select ren,llid from xmren where gid=" & Val(lblGid.Caption) & " order by llid desc"
            '    adoBlx.Close
            '    adoBlx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            
        '打开费用表
            'frmGzNr.dtgFy.Columns(0).Button = False
            tt = "Select * from fyTG where gid=" & Val(frmGzNr.lblGid.Caption)
            frmGzNr.adoFy.Recordset.Close
            frmGzNr.adoFy.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
            Set frmGzNr.dtgFy.DataSource = frmGzNr.adoFy
            
        
            
            
            
            If (mod1.KhK = 1 And frmGzNr.lblYwy <> mod1.DName) Or mod1.KhK = 2 Then
                frmGzNr.txtzgPd.Locked = False
            Else
                frmGzNr.txtzgPd.Locked = True
            End If
            'frmGzNr.frmMod.Enabled = True
            frmGzNr.cmdSave.Enabled = True
            frmGzNr.cmdDel.Enabled = False
            frmGzNr.cmdMod.Enabled = False
            frmGzNr.frmFy.Visible = False
            frmGzNr.Visible = True
            
            frmGzNr.cmdRenAdd.Visible = True
           frmGzNr.cmdRenDel.Visible = True
           frmGzNr.cmdFadd.Visible = True
           frmGzNr.cmdFdel.Visible = True
           frmGzNr.cmdTg.Visible = True
           frmGzNr.comRen.Visible = True
          ' frmGzNr.txtJw.Locked = True
        
        '设置流程按钮
        Call modXmGz.BGLcBut(40)


    

    
    End If

End If

End Sub

Private Sub cmdPre_Click()
Dim Ye As Integer
Ye = Year(dtpL.Value)
dtpL.Value = DateSerial(Year(dtpL.Value), Month(dtpL.Value) - 1, 1)
If Ye = Year(dtpL.Value) Then
    dtpR.Value = DateSerial(Year(dtpR.Value), Month(dtpL.Value) + 1, Day(dtpL.Value) - 1)
Else
    dtpR.Value = DateSerial(Year(dtpR.Value) - 1, Month(dtpL.Value) + 1, Day(dtpL.Value) - 1)
End If
End Sub

Private Sub cmdPz_Click()
Dim tt As String
Dim LLR As Date
On Error Resume Next
modXmGz.FR = DateSerial(Year(modXmGz.FR), Month(modXmGz.FR), Day(modXmGz.FR) - 7)
modXmGz.LR = DateSerial(Year(modXmGz.LR), Month(modXmGz.LR), Day(modXmGz.LR) - 7)
LLR = DateSerial(Year(modXmGz.LR), Month(modXmGz.LR), Day(modXmGz.LR) + 1)
frmGzBG.lblFR.Caption = modXmGz.FR
frmGzBG.lblLR.Caption = modXmGz.LR
tt = "Select atime,khqc,xmfy,NewF,gid from xmgz where ywy like '%" & lblYwy.Caption & "%' and aTime>='" & modXmGz.FR & _
"' and aTime <'" & LLR & "' and lb=1 order by aTime"
frmGzBG.adoXm.Close
frmGzBG.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText

Set frmGzBG.dtgXmgz.DataSource = frmGzBG.adoXm
tt = "Select atime,xmmc,newF,gid from xmgz where ywy like '%" & lblYwy.Caption & "%' and aTime>='" & modXmGz.FR & _
"' and aTime <='" & modXmGz.LR & "' and lb=0 order by aTime"
frmGzBG.adoJi.Close
frmGzBG.adoJi.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGzBG.dtgJi.DataSource = frmGzBG.adoJi
dtPRQ.Value = lblFR.Caption
lblWeek.Caption = modXmGz.dayWeek(dtPRQ.DayOfWeek)
End Sub

Private Sub cmdXZ_Click()
Set Ren.XForm = New frmGzBG
Call mod1.RenXz(Me.Name, Me, 0)
End Sub

Private Sub comKhmc_Click()
Dim tt As String
On Error Resume Next
If comKhmc.Text = "" Then Exit Sub

'待发布
tt = "Select 合同金额,合同编号 from htView where xywy='" & mod1.DName & "' and xuid='" & mod1.DHid & "' and 状态='执行' and 项目名称='" & comKhmc.Text & "' order by 合同日期 desc"
adoHT.Close
adoHT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgHt.DataSource = adoHT
lblHtbh.Caption = ""
End Sub

Private Sub comKhmc_DropDown()
Dim oo As Integer
Dim jj As Integer
Dim tt As String
On Error Resume Next


    '设置客户名称下拉框
    Set comKhmc.DataSource = Nothing
    jj = comKhmc.ListCount
    If jj > 0 Then
        For oo = jj - 1 To 0 Step -1
            comKhmc.RemoveItem (oo)
        Next
    End If
    'If mod1.KhK = 0 Then
        tt = "Select xmmc from xmzl where ywy='" & mod1.DName & "' and xmmc like '%" & _
        comKhmc.Text & "%'"
'    ElseIf mod1.KhK = 1 Then '外地经理,上海销售经理不填
'        tt = "Select khQc from khzl where  khQc like '%" & frmGzBG1.comKhmc.Text & "%'" & _
'        " and xmqy='" & mod1.Qy & "' group by khqc"
    'End If
    adoKhmc.Recordset.Close
    adoKhmc.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    adoKhmc.Recordset.MoveLast
    jj = adoKhmc.Recordset.RecordCount
    adoKhmc.Recordset.MoveFirst
    For oo = 0 To jj - 1
        comKhmc.AddItem adoKhmc.Recordset.Fields("xmmc").Value, oo
        adoKhmc.Recordset.MoveNext
    Next
   'comKhmc.AddItem "豪曼"

End Sub








Private Sub dtgHt_Click()
lblHtbh.Caption = adoHT.Fields("合同编号").Value
End Sub

Private Sub dtgJi_Click()
On Error Resume Next
'modXmGz.Gid = adoJi.Recordset.Fields("gid").Value
End Sub

Private Sub dtgJi_DblClick()
On Error Resume Next
frmGzJ.Show
Call modXmGz.jhQing
'Set frmGzJ.lblXmmc.DataSource = Nothing

Call modXmGz.jiBound

modXmGz.Ti = False
If mod1.KhK >= 1 Then
    frmGzJ.txtzgPd.Locked = False
Else
    frmGzJ.txtzgPd.Locked = True
End If
    frmGzJ.frmMod.Enabled = False
    frmGzJ.frmMod.Enabled = False
    frmGzJ.cmdSave.Enabled = False
'    If (frmZu.comYwy.Text = frmLogin.Combo1.Text And frmGzJ.txtzgPd.Text = "") Or mod1.ZW = "系统管理员" Or _
'    mod1.KhK >= 1 Then
'        frmGzJ.cmdDel.Enabled = True
'    Else
'        frmGzJ.cmdDel.Enabled = False
'    End If
    
End Sub

Private Sub dtgJi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
modXmGz.Gid = adoJi.Fields("gid").Value

End Sub

Private Sub dtgKH_DblClick()
Dim tt As String
Dim Xmmc As String
On Error Resume Next
dtgKh.Col = 1
Xmmc = dtgKh.Text
Set frmGzBG.dtgXmgz.DataSource = frmGzBG.adoXm
tt = "Select atime as 日期,xmmc as 项目名称,sum(xmfy) as 项目费用,gid,newF from xmgzB where ywy ='" & lblYwy.Caption & "' and xmmc='" & Xmmc & "' and atime>='" & _
 dtpL.Value & "' and atime<'" & DateSerial(Year(dtpR.Value), Month(dtpR.Value), Day(dtpR.Value) + 1) & "' group by atime,xmmc,gid,newf order by aTime"
frmGzBG.adoMx.Close
frmGzBG.adoMx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGzBG.dtgMx.DataSource = frmGzBG.adoMx
End Sub


Private Sub dtgMx_DblClick()
Dim Gid As Long
On Error Resume Next
dtgMx.Col = 4
Gid = dtgMx.Text
If Gid = 0 Or IsNull(Gid) = True Then
    Exit Sub
End If
frmGzNr.Show
Call modXmGz.xmQing
Set frmGzNr.lblXmmc.DataSource = Nothing
Call modXmGz.xmBound(Gid)

modXmGz.Ti = False

frmGzNr.txtzgPd.Locked = True



frmGzNr.cmdSave.Enabled = False
End Sub


Private Sub dtgXmgz_Click()
On Error Resume Next
modXmGz.Gid = adoXm.Fields("gid").Value
End Sub

Private Sub dtgXmgz_DblClick()
On Error Resume Next
frmGzNr.Show
Call modXmGz.xmQing
Set frmGzNr.lblXmmc.DataSource = Nothing
Call modXmGz.xmBound(adoXm.Fields("gid").Value)

modXmGz.Ti = False
If mod1.KhK = 1 Or mod1.KhK = 2 Then
    frmGzNr.txtzgPd.Locked = False
Else
    frmGzNr.txtzgPd.Locked = True
    
End If


    frmGzNr.cmdSave.Enabled = False
'    If (frmZu.comYwy.Text = frmLogin.Combo1.Text And frmGzNr.txtzgPd.Text = "") Or mod1.ZW = "系统管理员" Or _
'    mod1.ZW = "客户服务总监" Or mod1.ZW = "营销部总监" Then
'        frmGzNr.cmdDel.Enabled = True
'    Else
'        frmGzNr.cmdDel.Enabled = False
'    End If
    
End Sub


Private Sub dtgXmgz_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
modXmGz.Gid = adoXm.Fields("gid").Value
End Sub


Private Sub dtpRq_CloseUp()
lblWeek.Caption = modXmGz.dayWeek(dtPRQ.DayOfWeek)

End Sub


Private Sub dtpXr_CloseUp()
Dim tt As String
On Error Resume Next


Select Case dtpXr.DayOfWeek
Case 1 '星期日
modXmGz.FR = DateSerial(Year(dtpXr.Value), Month(dtpXr.Value), Day(dtpXr.Value) - 6)
modXmGz.LR = dtpXr.Value
Case 2 '星期一
modXmGz.FR = dtpXr.Value
modXmGz.LR = DateSerial(Year(dtpXr.Value), Month(dtpXr.Value), Day(dtpXr.Value) + 6)
Case 3
modXmGz.FR = DateSerial(Year(dtpXr.Value), Month(dtpXr.Value), Day(dtpXr.Value) - 1)
modXmGz.LR = DateSerial(Year(dtpXr.Value), Month(dtpXr.Value), Day(dtpXr.Value) + 5)
Case 4
modXmGz.FR = DateSerial(Year(dtpXr.Value), Month(dtpXr.Value), Day(dtpXr.Value) - 2)
modXmGz.LR = DateSerial(Year(dtpXr.Value), Month(dtpXr.Value), Day(dtpXr.Value) + 4)
Case 5
modXmGz.FR = DateSerial(Year(dtpXr.Value), Month(dtpXr.Value), Day(dtpXr.Value) - 3)
modXmGz.LR = DateSerial(Year(dtpXr.Value), Month(dtpXr.Value), Day(dtpXr.Value) + 3)
Case 6
modXmGz.FR = DateSerial(Year(dtpXr.Value), Month(dtpXr.Value), Day(dtpXr.Value) - 4)
modXmGz.LR = DateSerial(Year(dtpXr.Value), Month(dtpXr.Value), Day(dtpXr.Value) + 2)
Case 7
modXmGz.FR = DateSerial(Year(dtpXr.Value), Month(dtpXr.Value), Day(dtpXr.Value) - 5)
modXmGz.LR = DateSerial(Year(dtpXr.Value), Month(dtpXr.Value), Day(dtpXr.Value) + 1)
End Select

frmGzBG.lblFR.Caption = modXmGz.FR
frmGzBG.lblLR.Caption = modXmGz.LR


tt = "Select atime,khqc,xmfy,NewF,gid from xmgz where ywy like '%" & lblYwy.Caption & "%' and aTime>='" & modXmGz.FR & _
"' and aTime <='" & modXmGz.LR & "' and lb=1 order by aTime"
frmGzBG.adoXm.Close
frmGzBG.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText

Set frmGzBG.dtgXmgz.DataSource = frmGzBG.adoXm
tt = "Select atime,khqc,newF,gid from xmgz where ywy like '%" & lblYwy.Caption & "%' and aTime>='" & modXmGz.LR & _
"' and aTime <='" & modXmGz.LR & "' and lb=0 order by aTime"
frmGzBG.adoJi.Close
frmGzBG.adoJi.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGzBG.dtgJi.DataSource = frmGzBG.adoJi
dtPRQ.Value = lblFR.Caption
lblWeek.Caption = modXmGz.dayWeek(dtPRQ.DayOfWeek)

dtPRQ.Value = lblFR.Caption
lblWeek.Caption = modXmGz.dayWeek(dtPRQ.DayOfWeek)
End Sub


Private Sub Form_Load()
Set adoHT = CreateObject("adodb.recordset")
frmGzBG.Height = mod1.FHeight
frmGzBG.Width = mod1.FWidth
frmGzBG.Left = 0
frmGzBG.Top = 0
dtgKh.ColWidth(0) = 0
dtgKh.ColWidth(1) = 3000 '项目名称
dtgKh.ColWidth(2) = 500 '平台
dtgKh.ColWidth(3) = 800 '费用
dtgMx.ColWidth(0) = 0
dtgMx.ColWidth(2) = 2800
dtgMx.ColWidth(4) = 0
dtgMx.ColWidth(5) = 0
dtPRQ.Value = mod1.DQda
dtpXr.Value = mod1.DQda
Set adoXm = CreateObject("adodb.recordset")
Set adoJi = CreateObject("adodb.recordset")
Set adoRq = CreateObject("adodb.recordset")
Set adoMx = CreateObject("adodb.recordset")
dtpL.Value = DateSerial(Year(Date), Month(Date), 1)
dtpR.Value = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
'If mod1.DName = "谢雪梅" Then
'If mod1.DName = "谢雪梅" Or mod1.DName = "罗红盛" Then
    cmdNew.Visible = True
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
Cancel = True
frmZu.Enabled = True
End If
End Sub


