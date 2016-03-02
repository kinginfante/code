VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGGL 
   Caption         =   "公告栏"
   ClientHeight    =   7185
   ClientLeft      =   7695
   ClientTop       =   3015
   ClientWidth     =   10440
   ForeColor       =   &H00004080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmGGL.frx":0000
   ScaleHeight     =   7185
   ScaleWidth      =   10440
   Begin VB.PictureBox pic1 
      Height          =   795
      Left            =   630
      Picture         =   "frmGGL.frx":20CC1
      ScaleHeight     =   735
      ScaleWidth      =   975
      TabIndex        =   49
      Top             =   1950
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Timer timFl 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   3030
      Top             =   6420
   End
   Begin VB.Timer timNew 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   1950
      Top             =   6120
   End
   Begin VB.Frame frmCb 
      BackColor       =   &H00FFFFFF&
      Caption         =   "frmCb"
      Height          =   825
      Left            =   360
      TabIndex        =   38
      Top             =   6120
      Width           =   4665
      Begin VB.Label lblCf 
         BackStyle       =   0  'Transparent
         Caption         =   "再"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   810
         TabIndex        =   43
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblCg 
         BackStyle       =   0  'Transparent
         Caption         =   "做"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   1605
         TabIndex        =   42
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblCh 
         BackStyle       =   0  'Transparent
         Caption         =   "好"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   2400
         TabIndex        =   41
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblCi 
         BackStyle       =   0  'Transparent
         Caption         =   "事"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   3195
         TabIndex        =   40
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblCj 
         BackStyle       =   0  'Transparent
         Caption         =   "情"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   3990
         TabIndex        =   39
         Top             =   240
         Width           =   555
      End
   End
   Begin MSAdodcLib.Adodc adoFile 
      Height          =   375
      Left            =   5340
      Top             =   6750
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\HMXP9000\work.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\HMXP9000\work.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "htPing1"
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
   Begin MSComDlg.CommonDialog cmdDia 
      Left            =   3870
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmRen 
      Caption         =   "             选择发送人"
      Height          =   7185
      Left            =   7440
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   4515
      Begin VB.CommandButton cmdFA 
         BackColor       =   &H00FFC0C0&
         Caption         =   "在线全部发送"
         Height          =   285
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   6330
         Width           =   1215
      End
      Begin VB.CommandButton cmdRef 
         Caption         =   "刷新"
         Height          =   285
         Left            =   30
         TabIndex        =   47
         Top             =   6360
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgOline 
         Height          =   5925
         Left            =   30
         TabIndex        =   46
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   10451
         _Version        =   393216
         ForeColor       =   16711680
         FixedCols       =   0
         ForeColorFixed  =   255
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdNM 
         BackColor       =   &H00FF00FF&
         Caption         =   "匿名发送"
         Height          =   345
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   6750
         Width           =   885
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   1185
         Left            =   2190
         TabIndex        =   21
         Top             =   240
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   2090
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Hello!"
         TabPicture(0)   =   "frmGGL.frx":429B0
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cmdAll"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdXZ"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin VB.CommandButton cmdXZ 
            Caption         =   "选择人员"
            Height          =   285
            Left            =   210
            TabIndex        =   44
            Top             =   480
            Width           =   1905
         End
         Begin VB.CommandButton cmdAll 
            Caption         =   "豪曼集团所有人"
            Height          =   315
            Left            =   210
            TabIndex        =   22
            Top             =   120
            Width           =   1905
         End
      End
      Begin VB.CommandButton cmdRe 
         Caption         =   "重置"
         Height          =   255
         Left            =   2430
         TabIndex        =   20
         Top             =   6390
         Width           =   645
      End
      Begin VB.CommandButton cmdRd 
         Caption         =   "删除"
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   6390
         Width           =   645
      End
      Begin MSDataGridLib.DataGrid dtgRen 
         Height          =   4485
         Left            =   2310
         TabIndex        =   17
         Top             =   1560
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   7911
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   11.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   1
         BeginProperty Column00 
            DataField       =   "username"
            Caption         =   "姓名"
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
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdF 
         BackColor       =   &H00C0FFC0&
         Caption         =   "发  送"
         Height          =   315
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6780
         Width           =   1035
      End
   End
   Begin VB.Frame frmLx 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   345
      Left            =   7710
      TabIndex        =   28
      Top             =   840
      Width           =   1575
      Begin VB.ComboBox comLb 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   360
         ItemData        =   "frmGGL.frx":429CC
         Left            =   0
         List            =   "frmGGL.frx":429E2
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   0
         Width           =   1605
      End
   End
   Begin MSAdodcLib.Adodc adoRen 
      Height          =   330
      Left            =   3600
      Top             =   5520
      Visible         =   0   'False
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc adoGG 
      Height          =   435
      Left            =   4710
      Top             =   930
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   767
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4470
      Left            =   1470
      ScaleHeight     =   4440
      ScaleWidth      =   7785
      TabIndex        =   5
      Top             =   1290
      Width           =   7815
      Begin VB.Frame frmCa 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "frmCa"
         Height          =   615
         Left            =   390
         TabIndex        =   32
         Top             =   960
         Width           =   4785
         Begin VB.Label lblce 
            BackStyle       =   0  'Transparent
            Caption         =   "情"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   3360
            TabIndex        =   37
            Top             =   0
            Width           =   555
         End
         Begin VB.Label lblcd 
            BackStyle       =   0  'Transparent
            Caption         =   "心"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   2565
            TabIndex        =   36
            Top             =   0
            Width           =   555
         End
         Begin VB.Label lblcc 
            BackStyle       =   0  'Transparent
            Caption         =   "理"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   1770
            TabIndex        =   35
            Top             =   0
            Width           =   555
         End
         Begin VB.Label lblCb 
            BackStyle       =   0  'Transparent
            Caption         =   "处"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   975
            TabIndex        =   34
            Top             =   0
            Width           =   555
         End
         Begin VB.Label lblCa 
            BackStyle       =   0  'Transparent
            Caption         =   "先"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   180
            TabIndex        =   33
            Top             =   0
            Width           =   555
         End
      End
      Begin VB.CommandButton cmdXJ 
         BackColor       =   &H00C0FFC0&
         Caption         =   "已看"
         Height          =   315
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3960
         Width           =   675
      End
      Begin VB.CommandButton cmdBr 
         BackColor       =   &H00C0FFC0&
         Caption         =   "浏览"
         Height          =   315
         Left            =   6930
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3960
         Width           =   675
      End
      Begin VB.CommandButton cmdReply 
         BackColor       =   &H00C0FFC0&
         Caption         =   "回复"
         Height          =   315
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3960
         Width           =   675
      End
      Begin VB.CommandButton cmdYjb 
         BackColor       =   &H00C0C0FF&
         Caption         =   "查看投诉"
         Height          =   555
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3180
         Width           =   1275
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00C0FFC0&
         Caption         =   "后一条"
         Height          =   315
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3960
         Width           =   675
      End
      Begin VB.CommandButton cmdPre 
         BackColor       =   &H00C0FFC0&
         Caption         =   "前一条"
         Height          =   315
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3960
         Width           =   675
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0FFC0&
         Caption         =   "待发"
         Height          =   315
         Left            =   3570
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3960
         Width           =   675
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00C0FFC0&
         Caption         =   "删除"
         Height          =   315
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3960
         Width           =   675
      End
      Begin VB.CommandButton cmdZx 
         BackColor       =   &H00C0FFC0&
         Caption         =   "撰写"
         Height          =   315
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3960
         Width           =   675
      End
      Begin RichTextLib.RichTextBox rihNr 
         Height          =   3465
         Left            =   330
         TabIndex        =   6
         Top             =   420
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   6112
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmGGL.frx":42A1A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   1785
         Left            =   960
         Top             =   1950
         Width           =   1365
      End
      Begin VB.Label lblDate 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1260
         TabIndex        =   13
         Top             =   4020
         Width           =   945
      End
      Begin VB.Label lblZZ 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Height          =   285
         Left            =   330
         TabIndex        =   12
         Top             =   4020
         Width           =   735
      End
   End
   Begin VB.Timer Timer5 
      Interval        =   319
      Left            =   6360
      Top             =   510
   End
   Begin VB.Timer Timer4 
      Interval        =   319
      Left            =   5670
      Top             =   510
   End
   Begin VB.Timer Timer3 
      Interval        =   319
      Left            =   4800
      Top             =   390
   End
   Begin VB.Timer Timer2 
      Interval        =   319
      Left            =   3600
      Top             =   330
   End
   Begin VB.Timer Timer1 
      Interval        =   319
      Left            =   3930
      Top             =   90
   End
   Begin VB.OLE OLE2 
      Class           =   "Word.Document.8"
      Height          =   495
      Left            =   6390
      OleObjectBlob   =   "frmGGL.frx":42D58
      TabIndex        =   31
      Top             =   6000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      DataSource      =   "adoFile"
      Height          =   225
      Left            =   6840
      TabIndex        =   30
      Top             =   6060
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblLb 
      BackStyle       =   0  'Transparent
      Caption         =   "类别:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H002CFCF7&
      Height          =   255
      Left            =   6960
      TabIndex        =   27
      Top             =   930
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      DataField       =   "UserId"
      DataSource      =   "adoRen"
      Height          =   255
      Left            =   5790
      TabIndex        =   18
      Top             =   5550
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataField       =   "UserId"
      DataSource      =   "adoGG"
      Height          =   255
      Left            =   5490
      TabIndex        =   14
      Top             =   6270
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblE 
      BackStyle       =   0  'Transparent
      Caption         =   "春"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   270
      TabIndex        =   4
      Top             =   4560
      Width           =   555
   End
   Begin VB.Label lblD 
      BackStyle       =   0  'Transparent
      Caption         =   "新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   270
      TabIndex        =   3
      Top             =   3876
      Width           =   555
   End
   Begin VB.Label lblC 
      BackStyle       =   0  'Transparent
      Caption         =   "迎"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   270
      TabIndex        =   2
      Top             =   3194
      Width           =   555
   End
   Begin VB.Label lblB 
      BackStyle       =   0  'Transparent
      Caption         =   "啸"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   270
      TabIndex        =   1
      Top             =   2512
      Width           =   555
   End
   Begin VB.Label lblA 
      BackStyle       =   0  'Transparent
      Caption         =   "虎"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   270
      TabIndex        =   0
      Top             =   1830
      Width           =   555
   End
End
Attribute VB_Name = "frmGGL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public adoGGl As Object
Dim adoGRen As Object

  Const FLASHW_STOP = 0                                                                   'Stop   flashing.   The   system   restores _
                                                                                                          the   window   to   its   original   state.
  Const FLASHW_CAPTION = &H1                                                         'Flash   the   window   caption.
  Const FLASHW_TRAY = &H2                                                               'Flash   the   taskbar   button.
  Const FLASHW_ALL = (FLASHW_CAPTION Or FLASHW_TRAY)             'Flash   both   the   window   caption   and   taskbar   button.   This   is _
                                                                                                          equivalent   to   setting   the   FLASHW_CAPTION   Or   FLASHW_TRAY   flags.
  Const FLASHW_TIMER = &H4                                                             'Flash   continuously,   until   the   FLASHW_STOP   flag   is   set.
  Const FLASHW_TIMERNOFG = &HC                                                     'Flash   continuously   until   the   window   comes   to   the   foreground.
  Private Type FLASHWINFO
      cbSize         As Long
      hwnd             As Long
      dwFlags       As Long
      uCount         As Long
      dwTimeout   As Long
  End Type
  Private Declare Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Boolean
  Private Declare Sub Sleep Lib "kernel32" _
   (ByVal dwMilliseconds As Long)
   
  Public Gid As Long
  
  Public FlId As Long

Public FAll As Boolean '是否为全集团发送
Private Sub cmdAll_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
ii = MsgBox("您是否要发送给所有同事?(如果您的这条消息是发送给个人,则群发会导致您的公告栏收到垃圾信息!)", vbInformation + vbYesNo, "请确认您的操作!!!")
If ii = vbNo Then
    Exit Sub
End If
adoRen.Recordset.MoveFirst
Do While Not adoRen.Recordset.EOF
    adoRen.Recordset.Delete adAffectCurrent
    adoRen.Recordset.MoveNext
Loop

tt = "Select username from worker where yof=1 and zzF=1 order by username"
Set mod1.HTP = CreateObject("adodb.recordset")

mod1.HTP.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
Do While Not mod1.HTP.EOF
    adoRen.Recordset.AddNew "username", mod1.HTP.Fields("username").Value
    mod1.HTP.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub


Private Sub cmdBd_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username,userid from worker where qy='上海' and zzF=1"
adoGRen.Close
'基础发布
'Select Case mod1.Lqy
'Case "上海"
'    adoGRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    adoGRen.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
adoGRen.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
adoGRen.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not adoGRen.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = adoGRen.Fields("username").Value And adoRen.Recordset.Fields("userid").Value = adoGRen.Fields("userid").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", adoGRen.Fields("username").Value
    End If
    adoGRen.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub


Private Sub cmdBJ_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username,userid from worker where qy='北京' and zzF=1"
adoGRen.Close
'基础发布
'Select Case mod1.Lqy
'Case "上海"
'    adoGRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    adoGRen.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
adoGRen.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
adoGRen.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not adoGRen.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = adoGRen.Fields("username").Value And adoRen.Recordset.Fields("userid").Value = adoGRen.Fields("userid").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", adoGRen.Fields("username").Value
    End If
    adoGRen.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub

Private Sub cmdBr_Click()
Dim tt As String
Dim oo As Integer
On Error Resume Next
frmGGLKan.Show

tt = "select left(gnr,10)+'...' as 内容提要,zz as 发送人,rq as 发送日期,gid,lb as 类别," & mod1.DName & " AS 看过否 from ggl where not(" & mod1.DName & _
     " is null) and rq>='" & DateSerial(Year(Date), Month(Date), Day(Date) - 1) & "' and left(zz,1)<>'n' order by " & mod1.DName & ",gid desc"
Set frmGGLKan.AdoJl = CreateObject("adodb.recordset")
frmGGLKan.AdoJl.Close
frmGGLKan.AdoJl.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGGLKan.dtgJl.DataSource = frmGGLKan.AdoJl
frmGGLKan.GGGCCC = False



For oo = 6 To 0 Step -1
    frmGGLKan.comBj.RemoveItem oo
Next
For oo = 9 To 0 Step -1
    frmGGLKan.txtZ.RemoveItem oo
Next
frmGGLKan.dtpZ.Visible = False
    frmGGLKan.comBj.AddItem "="
    frmGGLKan.comBj.Text = "="
    frmGGLKan.txtZ.AddItem "公告类"
    frmGGLKan.txtZ.AddItem "一般类"
    frmGGLKan.txtZ.AddItem "通知类"
    frmGGLKan.txtZ.AddItem "派工类"
    frmGGLKan.txtZ.AddItem "到帐类"
    frmGGLKan.txtZ.AddItem "晨会类"
    frmGGLKan.txtZ.AddItem "胡萝卜"
    frmGGLKan.txtZ.AddItem "其它类"
    frmGGLKan.txtZ.AddItem "评审修改"
    frmGGLKan.txtZ.Text = "公告类"
End Sub

Private Sub cmdCx_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username from worker where userBm='产销部' and zzF=1 and yoF=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
mod1.HTP.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not mod1.HTP.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = mod1.HTP.Fields("username").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", mod1.HTP.Fields("username").Value
    End If
    mod1.HTP.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub

Private Sub cmdCj_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username from worker where userBm='产技部' and zzF=1 and yoF=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
mod1.HTP.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not mod1.HTP.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = mod1.HTP.Fields("username").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", mod1.HTP.Fields("username").Value
    End If
    mod1.HTP.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub


Private Sub cmdF_Click()
Dim tt As String
Dim oo As Integer
Dim Tnr As String
Dim Tzz As String
Dim Trq As Date
'Dim FaRen As String '发送人的字符串通一气
On Error Resume Next

If mod1.Mname = "马晓聪" Then
        tt = "declare @gid int;" & _
            "insert into Nggl (gnr,zuid,rq,lb,qf) values ('" & Left(rihNr.Text, 2000) & "','" & mod1.DHid & "',getdate(),'" & comLb.Text & "',255)"
    If FAll = True Then

    Else
        tt = tt & ";set @gid=@@identity"
        adoRen.Recordset.MoveFirst
        Do While Not adoRen.Recordset.EOF
            tt = tt & ";" & "insert into NgglDetail (gid,uid) values (@gid,'" & adoRen.Recordset.Fields("userid").Value & "')"
              adoRen.Recordset.MoveNext
        Loop
    End If
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.wzcc, adOpenForwardOnly, adLockReadOnly, adCmdText
    MsgBox "OK"
    FAll = False
    Exit Sub
End If

If Not (adoRen.Recordset.RecordCount = 1 And adoRen.Recordset.Fields(0).Value = "钱亘") And comLb.Text = "胡萝卜" Then
    MsgBox "胡萝卜的申请,必须指定为钱亘!"
    Exit Sub
End If
If Not (adoRen.Recordset.RecordCount = 1 And adoRen.Recordset.Fields(0).Value = "倪旭") And comLb.Text = "评审修改" Then
    MsgBox "评审修改申请,必须指定为倪旭!"
    Exit Sub
End If
rihNr.Locked = True
If rihNr.Text <> "" Then
    If adoRen.Recordset.RecordCount = 0 Then
        MsgBox "请选择发送人!"
        Exit Sub
    End If
    Tnr = rihNr.Text
    Tzz = mod1.DName
    Trq = mod1.DQda
    lblZZ.Caption = mod1.DName
    lblDate.Caption = mod1.DQda
    
'  If comRen.Text <> "所有人" Then
    '发布某个人
  tt = "Select top 0 * from ggl where gid=0"
  Set adoGGl = CreateObject("adodb.recordset")
  adoGGl.Close

  adoGGl.Open tt, mod1.workBh, adOpenKeyset, adLockBatchOptimistic, adCmdText
  adoGGl.AddNew "Gnr", Left(rihNr.Text, 2000)
  adoGGl.Update "zz", lblZZ.Caption
  adoGGl.Update "rq", lblDate.Caption
  adoGGl.Update "lb", comLb.Text
  adoGGl.Update "Fdx", frmGGL.rihNr.SelFontSize
  adoRen.Recordset.MoveFirst
  Do While Not adoRen.Recordset.EOF
        adoGGl.Fields(adoRen.Recordset.Fields("username").Value).Value = 0
        adoRen.Recordset.MoveNext
  Loop
  adoGGl.Fields(mod1.DName).Value = 1
On Error GoTo frmGGl_ErrC
  adoGGl.UpdateBatch
'    Call modGGL.GGLBound
'    rihNr.Text = Tnr
'    lblZZ.Caption = Tzz
'    lblDate.Caption = Trq
'adoRen.Recordset.MoveFirst
'Do While Not adoRen.Recordset.EOF
'    FaRen = adoRen.Recordset.Fields("username").Value & "=0"
'
'Loop
'tt = Insert
    On Error Resume Next
    If adoRen.Recordset.RecordCount = 1 Then
        MsgBox "您的消息已经发送给该同事了"
    Else
        MsgBox "您的消息已经公告天下 ：）"
    End If
'  Else


End If
cmdSave.Enabled = False
cmdZx.Enabled = True
cmdDel.Enabled = False
frmRen.Visible = False
Exit Sub
frmGGl_ErrC:
Call mod1.ErrInf
On Error Resume Next

End Sub

Private Sub cmdFx_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username from worker where userBm='风销部' and zzF=1 and yoF=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
mod1.HTP.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not mod1.HTP.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = mod1.HTP.Fields("username").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", mod1.HTP.Fields("username").Value
    End If
    mod1.HTP.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub

Private Sub cmdFA_Click()
Dim TG As Boolean
Dim oo As Integer
On Error Resume Next
oo = 1
For oo = 1 To Val(dtgOline.Row) + 1
    dtgOline.Row = oo
           frmGGL.adoRen.Recordset.MoveFirst
            TG = True
            Do While Not frmGGL.adoRen.Recordset.EOF
                If frmGGL.adoRen.Recordset.Fields("username").Value = dtgOline.Text Then
                    TG = False
                    Exit Do
                End If
                frmGGL.adoRen.Recordset.MoveNext
            Loop
            If TG = True Then
                dtgOline.Col = 0
                frmGGL.adoRen.Recordset.AddNew "username", dtgOline.Text
                dtgOline.Col = 1
                frmGGL.adoRen.Recordset.Update "userid", dtgOline.Text
                 Set frmGGL.dtgRen.DataSource = frmGGL.adoRen
            End If
Next
End Sub



Private Sub cmdGC_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username from worker where userBm='工程部' and zzF=1 and yoF=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
mod1.HTP.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not mod1.HTP.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = mod1.HTP.Fields("username").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", mod1.HTP.Fields("username").Value
    End If
    mod1.HTP.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub

Private Sub cmdGz_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username,userid from worker where qy='广州' and zzF=1"
adoGRen.Close
'基础发布
'Select Case mod1.Lqy
'Case "上海"
'    adoGRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    adoGRen.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
adoGRen.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
adoGRen.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not adoGRen.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = adoGRen.Fields("username").Value And adoRen.Recordset.Fields("userid").Value = adoGRen.Fields("userid").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", adoGRen.Fields("username").Value
    End If
    adoGRen.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub




Private Sub cmdHz_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username,userid from worker where qy='杭州' and zzF=1"
adoGRen.Close
'基础发布
'Select Case mod1.Lqy
'Case "上海"
'    adoGRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    adoGRen.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
adoGRen.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
adoGRen.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not adoGRen.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = adoGRen.Fields("username").Value And adoRen.Recordset.Fields("userid").Value = adoGRen.Fields("userid").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", adoGRen.Fields("username").Value
    End If
    adoGRen.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub

Private Sub cmdInput_Click()
Dim tt As String
Dim bt() As Byte
Dim Fid As Long
Dim Bm As String
cmdDia.ShowOpen
On Error Resume Next
If cmdDia.FileName = "" Then
    Exit Sub
End If
Bm = InputBox("请确认你的部门!(默认为" & mod1.Bm & ")")
If Bm = "" Then
    Bm = mod1.Bm
End If
Open cmdDia.FileName For Binary As #1
ReDim bt(LOF(1) - 1)
'ReDim bt(10000000)
    Get #1, , bt()
tt = "select * from gglfile where fid=0"
adoFile.Recordset.Close
adoFile.Recordset.Open tt, mod1.workBh, adOpenKeyset, adLockBatchOptimistic, adCmdText
adoFile.Recordset.AddNew "ywy", mod1.DName
adoFile.Recordset.Update "Fsize", LOF(1) - 1
adoFile.Recordset.Fields("FNR").AppendChunk bt()
adoFile.Recordset.UpdateBatch
'MsgBox adoFile.Recordset.Fields("fid").Value
Fid = adoFile.Recordset.Fields("fid").Value
'添加GGl表
  tt = "Select top 0 * from ggl where gid=0"
  Set adoGGl = CreateObject("adodb.recordset")
  adoGGl.Close
  adoGGl.Open tt, mod1.workBh, adOpenKeyset, adLockBatchOptimistic, adCmdText
  adoGGl.AddNew "Gnr", "早上好,今天,由我们" & Bm & "向您提供一篇精彩的晨会文章!"
  adoGGl.Update "zz", mod1.DName
  adoGGl.Update "rq", mod1.DQda
  adoGGl.Update "lb", comLb.Text
  adoGGl.Update "Fdx", frmGGL.rihNr.SelFontSize
  adoGGl.Update "fid", Fid
  adoRen.Recordset.MoveFirst
  Do While Not adoRen.Recordset.EOF
        adoGGl.Fields(adoRen.Recordset.Fields("username").Value).Value = 0
        adoRen.Recordset.MoveNext
  Loop
  adoGGl.Fields(mod1.DName).Value = 1
'On Error GoTo frmGGl_ErrC
  adoGGl.UpdateBatch
  MsgBox "您的精彩文章已经导入,将被全公司员工分享! :)"
End Sub

Private Sub cmdKx_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username from worker where userBm='空销部' and zzF=1 and yoF=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
mod1.HTP.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not mod1.HTP.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = mod1.HTP.Fields("username").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", mod1.HTP.Fields("username").Value
    End If
    mod1.HTP.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub

Private Sub cmdMa_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username from worker where userBm='总经理' and zzF=1 and yoF=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
mod1.HTP.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not mod1.HTP.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = mod1.HTP.Fields("username").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", mod1.HTP.Fields("username").Value
    End If
    mod1.HTP.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub

Private Sub cmdNext_Click()


MDI.ztT.Panels(3).Text = ""
Call mod1.YJJL '意见交流
If Timer1.Enabled = True Then
Call modGGL.zTing
End If
Call modGGL.CHZT '晨会彩字停
Call modGGL.GGLR

'Dim tt As String
'On Error Resume Next
'If Timer1.Enabled = True Then
'Call modGGL.zTing
'End If
'rihNr.Locked = True
'
'cmdPre.Enabled = True
'adoGG.Recordset.MoveNext
'adoGG.Recordset.MoveNext
'If adoGG.Recordset.EOF = True Then
'cmdNext.Enabled = False
'End If
'adoGG.Recordset.MovePrevious
'rihNr.Text = adoGG.Recordset.Fields("Gnr").Value
'
'lblZZ.Caption = adoGG.Recordset.Fields("zz").Value
'lblDate.Caption = adoGG.Recordset.Fields("rq").Value
'rihNr.SelStart = 0
'rihNr.SelLength = Len(rihNr.Text)
'If adoGG.Recordset.Fields(frmLogin.Combo1.Text).Value = 0 Then
'rihNr.SelColor = &HFF0000
'Else
'rihNr.SelColor = &H80000012
'End If
'rihNr.SelFontSize = adoGG.Recordset.Fields("Fdx").Value
'rihNr.Refresh
'rihNr.SelStart = 0
'rihNr.SelLength = 0
'
'If lblZZ.Caption = frmLogin.Combo1.Text Or frmLogin.lblZw.Caption = "系统管理员" Then
'cmdDel.Enabled = True
'Else
'cmdDel.Enabled = False
'End If
'cmdSave.Enabled = False
'frmGGL.cmdXQ.Visible = False
'frmGGL.cmdYjb.Visible = False
'If IsNull(frmGGL.adoGG.Recordset.Fields("wzid").Value) = False Then
'
'
'    If Left(frmGGL.rihNr.Text, 3) = "请注意" Then
'        frmGGL.cmdYjb.Visible = True
'    Else
'        frmGGL.cmdXQ.Visible = True
'    End If
'End If
'cmdZx.Enabled = True
'frmRen.Visible = False
'''''If mod1.ZT = "HMData" And frmTip.Visible = True Then
'''''    frmTip.Show
'''''    frmTip.ZOrder 0
'''''End If
End Sub

Private Sub cmdNj_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username,userid from worker where qy='南京' and zzF=1"
adoGRen.Close
'基础发布
'Select Case mod1.Lqy
'Case "上海"
'    adoGRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    adoGRen.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
adoGRen.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
adoGRen.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not adoGRen.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = adoGRen.Fields("username").Value And adoRen.Recordset.Fields("userid").Value = adoGRen.Fields("userid").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", adoGRen.Fields("username").Value
    End If
    adoGRen.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub


Private Sub cmdNM_Click()
Dim tt As String
Dim oo As Integer
Dim Tnr As String

Dim Trq As Date
'Dim FaRen As String '发送人的字符串
On Error Resume Next


rihNr.Locked = True
If rihNr.Text <> "" Then
    If adoRen.Recordset.RecordCount = 0 Then
        MsgBox "请选择发送人!"
        Exit Sub
    End If
    Tnr = rihNr.Text

    Trq = mod1.DQda
    lblZZ.Caption = "匿名者"
    lblDate.Caption = mod1.DQda
    
'  If comRen.Text <> "所有人" Then
    '发布某个人
  tt = "Select top 0 * from ggl where gid=0"
  Set adoGGl = CreateObject("adodb.recordset")
  adoGGl.Close

  adoGGl.Open tt, mod1.workBh, adOpenKeyset, adLockBatchOptimistic, adCmdText
  adoGGl.AddNew "Gnr", Left(rihNr.Text, 2000)
  adoGGl.Update "zz", "n" & mod1.DName
  adoGGl.Update "rq", lblDate.Caption
  adoGGl.Update "lb", comLb.Text
  adoGGl.Update "Fdx", frmGGL.rihNr.SelFontSize
  adoRen.Recordset.MoveFirst
  Do While Not adoRen.Recordset.EOF
        adoGGl.Fields(adoRen.Recordset.Fields("username").Value).Value = 0
        adoRen.Recordset.MoveNext
  Loop
  adoGGl.Fields(mod1.DName).Value = 1
On Error GoTo frmGGl_ErrC
  adoGGl.UpdateBatch
'    Call modGGL.GGLBound
'    rihNr.Text = Tnr
'    lblZZ.Caption = Tzz
'    lblDate.Caption = Trq
'adoRen.Recordset.MoveFirst
'Do While Not adoRen.Recordset.EOF
'    FaRen = adoRen.Recordset.Fields("username").Value & "=0"
'
'Loop
'tt = Insert
    On Error Resume Next
'  Else
    MsgBox "你的消息已经匿名发送,豪曼信息将保护您的隐私!"

End If
cmdSave.Enabled = False
cmdZx.Enabled = True
cmdDel.Enabled = False
frmRen.Visible = False
Exit Sub
frmGGl_ErrC:
Call mod1.ErrInf
On Error Resume Next
End Sub

Private Sub cmdPre_Click()
'On Error Resume Next
'If Timer1.Enabled = True Then
'Call modGGL.zTing
'End If
'rihNr.Locked = True
'cmdNext.Enabled = True
'adoGG.Recordset.MovePrevious
'adoGG.Recordset.MovePrevious
'If adoGG.Recordset.BOF = True Then
'cmdPre.Enabled = False
'End If
'adoGG.Recordset.MoveNext
'
'rihNr.Text = adoGG.Recordset.Fields("Gnr").Value
'lblZZ.Caption = adoGG.Recordset.Fields("zz").Value
'lblDate.Caption = adoGG.Recordset.Fields("rq").Value
'
'rihNr.SelStart = 0
'rihNr.SelLength = Len(rihNr.Text)
'If adoGG.Recordset.Fields(frmLogin.Combo1.Text).Value = 0 Then
'rihNr.SelColor = &HFF0000
'Else
'rihNr.SelColor = &H80000012
'End If
'rihNr.SelFontSize = adoGG.Recordset.Fields("Fdx").Value
'rihNr.SelStart = 0
'rihNr.SelLength = 0
'
'If lblZZ.Caption = frmLogin.Combo1.Text Or frmLogin.lblZw.Caption = "系统管理员" Then
'cmdDel.Enabled = True
'Else
'cmdDel.Enabled = False
'End If
'cmdSave.Enabled = False
'
'frmGGL.cmdXQ.Visible = False
'frmGGL.cmdYjb.Visible = False
'If IsNull(frmGGL.adoGG.Recordset.Fields("wzid").Value) = False Then
'
'
'    If Left(frmGGL.rihNr.Text, 3) = "请注意" Then
'        frmGGL.cmdYjb.Visible = True
'    Else
'        frmGGL.cmdXQ.Visible = True
'    End If
'End If
'cmdZx.Enabled = True
'frmRen.Visible = False


If Timer1.Enabled = True Then
Call modGGL.zTing
End If
Call modGGL.CHZT
Call modGGL.GGLL
End Sub



Private Sub cmdRd_Click()
On Error Resume Next

    adoRen.Recordset.Delete adAffectCurrent
    'Set dtgRen.DataSource = adoRen

End Sub

Private Sub cmdRe_Click()
On Error Resume Next
Dim tt As String
'adoRen.Recordset.MoveFirst
'Do While Not adoRen.Recordset.EOF
'    adoRen.Recordset.Delete adAffectCurrent
'    adoRen.Recordset.MoveNext
'Loop
tt = "Select top 1 username,userid from worker where qy='KKK'"
frmGGL.adoRen.Recordset.Close
frmGGL.adoRen.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Set dtgRen.DataSource = adoRen
End Sub

Private Sub cmdRef_Click()
Dim Ra
Dim La
Dim tt As String
Dim oo As Integer
On Error Resume Next
    tt = "select username,userid from oline order by bmid"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2)
    dtgOline.Rows = 30
    dtgOline.Col = 0
    dtgOline.Row = 1
    For oo = 1 To La + 1
        dtgOline.Row = oo
        dtgOline.Col = 0
        dtgOline.Text = Ra(0, oo - 1)
        dtgOline.Col = 1
        dtgOline.Text = Ra(1, oo - 1)
    Next
End Sub

Private Sub cmdReply_Click()
Dim ii As Integer
Dim tt As String
On Error Resume Next
If lblZZ.Caption = "匿名者" Then
    Exit Sub
End If
If Timer1.Enabled = True Then
Call modGGL.zTing
End If
If lblZZ.Caption = "匿名者" Then Exit Sub
If comLb.Text = "胡萝卜" And mod1.DName = "钱亘" Then
   ii = MsgBox("是否同意将此胡萝卜申请发出?", vbInformation + vbYesNo, "蛋蛋")
   If ii = vbYes Then
        Call cmdAll_Click
          tt = "Select top 0 * from ggl where gid=0"
          Set adoGGl = CreateObject("adodb.recordset")
          adoGGl.Close
        
          adoGGl.Open tt, mod1.workBh, adOpenKeyset, adLockBatchOptimistic, adCmdText
          adoGGl.AddNew "Gnr", "申请人:" & lblZZ.Caption & Chr(13) & Left(rihNr.Text, 2000)
          adoGGl.Update "zz", "钱亘"
          adoGGl.Update "rq", lblDate.Caption
          adoGGl.Update "lb", comLb.Text
          adoGGl.Update "Fdx", frmGGL.rihNr.SelFontSize
          adoRen.Recordset.MoveFirst
          Do While Not adoRen.Recordset.EOF
                adoGGl.Fields(adoRen.Recordset.Fields("username").Value).Value = 0
                adoRen.Recordset.MoveNext
          Loop
          adoGGl.Fields(mod1.DName).Value = 1
        On Error GoTo frmGGl_ErrC
          adoGGl.UpdateBatch
          MsgBox "胡萝卜通告已经发出!"
   Else
        comLb.Text = "一般类"
        rihNr.Text = ""
   End If
Else
    comLb.Text = "一般类"
    rihNr.Text = ""
End If
comLb.Visible = True
lblLb.Visible = True
comLb.Locked = True
frmLx.Enabled = False

lblZZ.Tag = lblZZ.Caption
lblZZ.Caption = ""
lblDate.Caption = ""
rihNr.Locked = False
cmdSave.Enabled = True
cmdDel.Enabled = False
frmGGL.rihNr.SelFontSize = 12

cmdReply.Enabled = False
cmdZx.Enabled = False
Exit Sub
frmGGl_ErrC:
End Sub

Private Sub cmdSave_Click()
Dim tt As String
Dim Ra
Dim La
Dim oo As Integer
On Error Resume Next
If comLb.Text = "请选择类别" Then
    MsgBox "请选择类别"
    Exit Sub
End If
If comLb.Text = "一般类" Then
    cmdAll.Enabled = False
ElseIf comLb.Text = "胡萝卜" And mod1.DName <> "钱亘" Then '员工向钱亘申请胡萝卜
    adoRen.Recordset.AddNew "username", "钱亘"
    Set dtgRen.DataSource = adoRen
ElseIf comLb.Text = "评审修改" Then
    adoRen.Recordset.AddNew "username", "倪旭"
    Set dtgRen.DataSource = adoRen
ElseIf comLb.Text = "胡萝卜" And mod1.DName = "钱亘" Then '钱亘向员工普发胡萝卜
    Call cmdAll_Click
Else
    cmdAll.Enabled = True
End If
If cmdReply.Enabled = True Then
    comLb.Locked = True
    frmLx.Enabled = False
    frmRen.Visible = True
    frmRen.Left = 5070
    frmRen.Top = 0
    tt = "select username,userid from oline order by bmid"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2)
    dtgOline.Rows = 30
    dtgOline.Col = 0
    dtgOline.Row = 1
    For oo = 1 To La + 1
        dtgOline.Row = oo
        dtgOline.Col = 0
        dtgOline.Text = Ra(0, oo - 1)
        dtgOline.Col = 1
        dtgOline.Text = Ra(1, oo - 1)
    Next
'comRen.Text = "所有人"
Else
  tt = "Select top 0 * from ggl where gid=0"
  adoGGl.Close
  adoGGl.Open tt, mod1.workBh, adOpenKeyset, adLockBatchOptimistic, adCmdText
  adoGGl.AddNew "Gnr", Left(rihNr.Text, 2000)
  adoGGl.Update "zz", mod1.DName
  adoGGl.Update "rq", mod1.DQda
  adoGGl.Update "Fdx", frmGGL.rihNr.SelFontSize
  adoGGl.Update "lb", comLb.Text
  adoGGl.Fields(lblZZ.Tag).Value = 0
  adoGGl.Fields(mod1.DName).Value = 1
  On Error GoTo frmGGl_ErrB
  adoGGl.UpdateBatch
'    Call modGGL.GGLBound
'    rihNr.Text = Tnr
'    lblZZ.Caption = Tzz
'    lblDate.Caption = Trq
'adoRen.Recordset.MoveFirst
'Do While Not adoRen.Recordset.EOF
'    FaRen = adoRen.Recordset.Fields("username").Value & "=0"
'
'Loop
lblZZ.Caption = mod1.DName
lblDate.Caption = mod1.DQda

        MsgBox "您的消息已经发送给" & lblZZ.Tag & " ：）"

End If

Exit Sub
frmGGl_ErrB:
Call mod1.ErrInf
End Sub

Private Sub cmdWe_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username from worker where userBm='维销部' and zzF=1 and yoF=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
mod1.HTP.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not mod1.HTP.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = mod1.HTP.Fields("username").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", mod1.HTP.Fields("username").Value
    End If
    mod1.HTP.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub

Private Sub cmdWX_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next

tt = "Select username,userid from worker where qy='无锡' and zzF=1"
adoGRen.Close
adoGRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
adoGRen.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not adoGRen.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = adoGRen.Fields("username").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", adoGRen.Fields("username").Value
    End If
    adoGRen.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub

Private Sub cmdXJ_Click()
If cmdXJ.Caption = "未看" Then
    cmdXJ.Caption = "已看"
    
Else
    cmdXJ.Caption = "未看"
End If
modGGL.Oid = 99999
Call modGGL.GGLR
cmdPre.Enabled = False
cmdNext.Enabled = True
End Sub

''''''''Private Sub cmdXQ_Click()
''''''''Dim bt() As Byte
''''''''Dim tt As String
''''''''On Error Resume Next
'''''''''OLE1.Close
'''''''''OLE2.Close
''''''''tt = "select fnr,fsize from gglfile where fid=" & cmdXQ.ToolTipText
''''''''adoFile.Recordset.Close
''''''''adoFile.Recordset.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
''''''''ReDim bt(adoFile.Recordset.Fields("Fsize").Value) As Byte
''''''''bt() = adoFile.Recordset.Fields("FNR").GetChunk(adoFile.Recordset.Fields("Fsize").Value + 1)
'''''''''bt() = adoFile.Recordset.Fields("FNR").GetChunk(1000000)
'''''''''FName = adoFile.Recordset.Fields("Fname").Value
'''''''''Set fs = CreateObject("Scripting.FileSystemObject")
'''''''''fs.deletefile ("c:\work\demo\file\" & Fname)
''''''''Open ("c:\work\demo\hmxp9000\" & "晨会文章") For Binary As #2
''''''''Put #2, , bt()
''''''''Close #2
''''''''
'''''''''tt = "Select * from hmfile where ywy='" & frmLogin.Combo1.Text & "'"
'''''''''frmFile.adoFile.Recordset.Close
'''''''''frmFile.adoFile.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'''''''''Set frmFile.dtGGF.DataSource = frmFile.adoFile
''''''''''判断打开类型
'''''''''If adoFile.Recordset.Fields("Flx").Value = "WORD" Then
''''''''    OLE2.SourceDoc = "c:\work\demo\hmxp9000\" & "晨会文章"
''''''''    OLE2.Action = 1
''''''''    OLE2.DoVerb (-2)
''''''''
'''''''''ElseIf adoFile.Recordset.Fields("Flx").Value = "EXCEL" Then
'''''''''    OLE1.SourceDoc = "c:\work\demo\file\" & FName
'''''''''    OLE1.Action = 1
'''''''''    OLE1.DoVerb (-2)
'''''''''End If
''''''''End Sub




Private Sub cmdYz_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username from worker where userBm='运作部' and zzF=1 and yoF=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
mod1.HTP.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not mod1.HTP.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = mod1.HTP.Fields("username").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", mod1.HTP.Fields("username").Value
    End If
    mod1.HTP.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub

Private Sub cmdXZ_Click()
Set Ren.XForm = New frmGGL
Call mod1.RenXz("frmGGL", Me, 0)
End Sub

Private Sub cmdZC_Click()
Dim TG As Boolean
Dim tt As String
On Error Resume Next
tt = "Select username from worker where zzF=1 and yoF=1 and zcg=1"
Set mod1.HTP = CreateObject("adodb.recordset")

mod1.HTP.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
mod1.HTP.MoveFirst
adoRen.Recordset.MoveFirst
Do While Not mod1.HTP.EOF
    TG = True
    adoRen.Recordset.MoveFirst
    Do While Not adoRen.Recordset.EOF
        If adoRen.Recordset.Fields("username").Value = mod1.HTP.Fields("username").Value Then
            TG = False
            Exit Do
        End If
        adoRen.Recordset.MoveNext
    Loop
    If TG = True Then
        adoRen.Recordset.AddNew "username", mod1.HTP.Fields("username").Value
    End If
    mod1.HTP.MoveNext
Loop
Set dtgRen.DataSource = adoRen
End Sub

Private Sub cmdZX_Click()
If Timer1.Enabled = True Then
Call modGGL.zTing
End If
rihNr.Text = ""
lblZZ.Caption = ""
lblDate.Caption = ""
rihNr.Locked = False
cmdSave.Enabled = True
cmdZx.Enabled = False
cmdDel.Enabled = False
frmGGL.rihNr.SelFontSize = 12

cmdReply.Enabled = True
comLb.Locked = False
frmLx.Enabled = True
comLb.Visible = True
lblLb.Visible = True
comLb.Text = "一般类"
frmCa.Visible = False
frmCb.Visible = False

End Sub





Private Sub Command2_Click()

End Sub

Private Sub comLb_Click()

If comLb.Text = "公告类" Or comLb.Text = "一般类" Then
    cmdNM.Visible = True
Else
    cmdNM.Visible = False
End If
End Sub






Private Sub dtgOline_DblClick()
Dim TG As Boolean
On Error Resume Next
If dtgOline.Text <> "在线网友" And dtgOline.Text <> "" Then
           frmGGL.adoRen.Recordset.MoveFirst
            TG = True
            Do While Not frmGGL.adoRen.Recordset.EOF
                If frmGGL.adoRen.Recordset.Fields("username").Value = dtgOline.Text Then
                    TG = False
                    Exit Do
                End If
                frmGGL.adoRen.Recordset.MoveNext
            Loop
            If TG = True Then
                dtgOline.Col = 0
                frmGGL.adoRen.Recordset.AddNew "username", dtgOline.Text
                dtgOline.Col = 1
                frmGGL.adoRen.Recordset.Update "userid", dtgOline.Text
                 Set frmGGL.dtgRen.DataSource = frmGGL.adoRen
            End If
End If
End Sub


Private Sub Form_Click()
If Timer1.Enabled = True Then
Call modGGL.zTing
End If

MDI.ztT.Panels(3).Text = ""

End Sub

Private Sub Form_GotFocus()
MDI.ztT.Panels(3).Text = ""
End Sub

Private Sub Form_Load()
Dim tt As String
Dim oo As Integer
On Error Resume Next

frmGGL.Width = 9720
frmGGL.Height = 7590
'''''frmHz.Left = 30
'''''frmHz.Top = 0
frmGGL.Left = (Screen.Width - frmGGL.Width) / 2
Set adoGRen = CreateObject("adodb.recordset")

'Me.Picture = LoadPicture(App.Path & "\pic\公告栏New.jpg")
Picture1.Picture = LoadPicture("c:\work\demo\hmxp9000\pic\Pan.jpg")
''判断字颜色
'rihNr.SelStart = 0
'rihNr.SelLength = Len(rihNr.Text)
'If adoGG.Recordset.Fields(frmLogin.Combo1.Text).Value = 0 Then
'rihNr.SelColor = &HFF0000
'Else
'rihNr.SelColor = &H80000012
'End If
'rihNr.SelFontSize = adoGG.Recordset.Fields("Fdx").Value
'rihNr.SelStart = 0
'rihNr.SelLength = 0
'If lblZZ.Caption = frmLogin.Combo1.Text Or frmLogin.lblZw.Caption = "系统管理员" Then
'cmdDel.Enabled = True
'Else
'cmdDel.Enabled = False
'End If
'cmdSave.Enabled = False
frmCa.BorderStyle = 0
frmCb.BorderStyle = 0
frmCa.Visible = False
frmCb.Visible = False

dtgOline.ColWidth(1) = 0
dtgOline.ColWidth(0) = 2000
dtgOline.Row = 0
dtgOline.Text = "在线网友"
End Sub






















Private Sub Form_Resize()
If Me.WindowState = 0 Then
    timFl.Enabled = False
    If Me.Visible = True Then
    modGGL.Oid = frmGGL.Gid
    End If
        'timNew.Enabled = False

End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

frmGGL.Visible = False
Cancel = True
frmZu.Enabled = True
frmGGLKan.Visible = False
frmZu.TBa.Buttons(2).Value = tbrUnpressed
End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub rihNr_Change()
If Len(rihNr.Text) >= 3000 Then
    MsgBox ("字数超过限制,超过部分将不被保存!")
End If
If comLb.Text = "请选择类别" Then
    MsgBox "请先选择类别!"
    rihNr.Text = ""
End If

End Sub

Private Sub rihNr_Click()
'Dim Tnr As String
On Error Resume Next
If Timer1.Enabled = True Then
Call modGGL.zTing
End If
'Tnr = rihNr.Text
'rihNr.SelText = Tnr
MDI.ztT.Panels(3).Text = ""
End Sub

Private Sub rihNr_KeyDown(KeyCode As Integer, Shift As Integer)

If Shift = 2 And KeyCode = 187 Then
    rihNr.SelStart = 0
    rihNr.SelLength = Len(rihNr.Text)
    rihNr.SelFontSize = rihNr.SelFontSize + 2
    If rihNr.SelFontSize > 30 Then
        rihNr.SelFontSize = 30
    End If
    rihNr.SelStart = 0
    rihNr.SelLength = 0
ElseIf Shift = 2 And KeyCode = 189 Then
    rihNr.SelStart = 0
    rihNr.SelLength = Len(rihNr.Text)
    rihNr.SelFontSize = rihNr.SelFontSize - 2
    If rihNr.SelFontSize < 8 Then
        rihNr.SelFontSize = 8
    End If
    rihNr.SelStart = 0
    rihNr.SelLength = 0
End If
End Sub


Private Sub rihNr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tt As String
On Error Resume Next
    If cmdZx.Enabled = True And cmdReply.Enabled = True And Button = 2 Then
    tt = "update ggl set " & mod1.DName & "=1 where gid=" & modGGL.Oid
    Set mod1.HTP = CreateObject("adodb.recordset")
    On Error GoTo frmggl_ErrA
    mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'mod1.HTP.Close
    rihNr.SelStart = 0
    rihNr.SelLength = Len(rihNr.Text)
    rihNr.SelColor = &H80000012
    rihNr.SelStart = 0
    rihNr.SelLength = 0

    End If
Exit Sub
frmggl_ErrA:
Call mod1.ErrInf
End Sub

Private Sub Timer1_Timer()
Dim X As Integer
Dim Y As Integer
Dim z As Integer
    X = Int(Rnd * (155)) + 100
    Y = Int(Rnd * (155)) + 100
    z = Int(Rnd * (155)) + 100
   lblA.ForeColor = RGB(X, Y, z)

End Sub

Private Sub Timer2_Timer()
Dim X As Integer
Dim Y As Integer
Dim z As Integer
    X = Int(Rnd * (155)) + 100
    Y = Int(Rnd * (155)) + 100
    z = Int(Rnd * (155)) + 100
   lblB.ForeColor = RGB(X, Y, z)
If comLb.Text = "晨会类" Then
    lblCb.ForeColor = RGB(X, Y, z)
    X = Int(Rnd * (155)) + 100
    Y = Int(Rnd * (155)) + 100
    z = Int(Rnd * (155)) + 100
    lblCg.ForeColor = RGB(X, Y, z)
    Timer2.Interval = 1000
End If
End Sub

Private Sub Timer3_Timer()
Dim X As Integer
Dim Y As Integer
Dim z As Integer
  X = Int(Rnd * (155)) + 100
    Y = Int(Rnd * (155)) + 100
    z = Int(Rnd * (155)) + 100
   lblC.ForeColor = RGB(X, Y, z)
If comLb.Text = "晨会类" Then
    lblcc.ForeColor = RGB(X, Y, z)
    X = Int(Rnd * (155)) + 100
    Y = Int(Rnd * (155)) + 100
    z = Int(Rnd * (155)) + 100
    lblCh.ForeColor = RGB(X, Y, z)
    Timer3.Interval = 1000
End If
End Sub

Private Sub Timer4_Timer()
Dim X As Integer
Dim Y As Integer
Dim z As Integer
  X = Int(Rnd * (155)) + 100
    Y = Int(Rnd * (155)) + 100
    z = Int(Rnd * (155)) + 100
   lblD.ForeColor = RGB(X, Y, z)
If comLb.Text = "晨会类" Then
    lblcd.ForeColor = RGB(X, Y, z)
    X = Int(Rnd * (155)) + 100
    Y = Int(Rnd * (155)) + 100
    z = Int(Rnd * (155)) + 100
    lblCi.ForeColor = RGB(X, Y, z)
    Timer4.Interval = 1000
End If
End Sub

Private Sub Timer5_Timer()
Dim X As Integer
Dim Y As Integer
Dim z As Integer
  X = Int(Rnd * (155)) + 100
    Y = Int(Rnd * (155)) + 100
    z = Int(Rnd * (155)) + 100
   lblE.ForeColor = RGB(X, Y, z)
If comLb.Text = "晨会类" Then
    lblce.ForeColor = RGB(X, Y, z)
    X = Int(Rnd * (155)) + 100
    Y = Int(Rnd * (155)) + 100
    z = Int(Rnd * (155)) + 100
    lblCj.ForeColor = RGB(X, Y, z)
    Timer5.Interval = 1000
End If
End Sub

Private Sub timFl_Timer()
      Dim FlashInfo     As FLASHWINFO

            'Specifies   the   size   of   the   structure.
            FlashInfo.cbSize = Len(FlashInfo)
            'Specifies   the   flash   status
            FlashInfo.dwFlags = FLASHW_ALL Or FLASHW_TIMER
            'Specifies   the   rate,   in   milliseconds,   at   which   the   window   will   be   flashed.   If _
              dwTimeout   is   zero,   the   function   uses   the   default   cursor   blink   rate.
            FlashInfo.dwTimeout = 0
            'Handle   to   the   window   to   be   flashed.   The   window   can   be   either   opened   or   minimized.
            FlashInfo.hwnd = Me.hwnd
            'Specifies   the   number   of   times   to   flash   the   window.
            FlashInfo.uCount = 1
            FlashWindowEx FlashInfo

End Sub

Private Sub timNew_Timer()
Dim tt As String

Dim zz As String
Dim Ra
Dim La

On Error Resume Next
Set frmGGL.adoGGl = CreateObject("adodb.recordset")
        tt = "Select gnr,zz,rq,gid,fdx,wzid,lb,fid  from ggl where  gid>" & FlId & " and  (" & mod1.DName & "=0 or " & mod1.DName & " is null and lb='胡萝卜') order by gid desc;" & _
            "update HMDATA.dbo.worker set Oline=getdate(),cname='" & mod1.CName & "' where userid='" & mod1.DHid & "'"
    frmGGL.adoGGl.Open tt, mod1.workBh, adOpenForwardOnly, adLockReadOnly, adCmdText
  Ra = frmGGL.adoGGl.GetRows
  La = UBound(Ra, 2) + 1
  Gid = Ra(3, 0): zz = Ra(1, 0)
  If Left(zz, 1) = "n" Then
    zz = "匿名者"
  End If
  
  If La = 1 Then
      If cmdZx.Enabled = True Then
  

                On Error Resume Next

            

            
                    modGGL.Oid = Ra(3, 0)
                    FlId = Ra(3, 0)
                    frmGGL.rihNr.Text = Ra(0, 0)
            
                    frmGGL.lblZZ.Caption = zz
            
                
                    frmGGL.lblDate.Caption = Ra(2, 0)
            
                        frmGGL.comLb.Text = Ra(6, 0)
                        frmGGL.comLb.Visible = True
                        frmGGL.lblLb.Visible = True
                        frmGGL.comLb.Locked = True
                        frmGGL.frmLx.Enabled = False

            
                
                    '判断字颜色
                    frmGGL.rihNr.SelStart = 0
                    frmGGL.rihNr.SelLength = Len(frmGGL.rihNr.Text)
                
                        frmGGL.rihNr.SelColor = &HFF0000
                
                    frmGGL.rihNr.SelFontSize = Ra(4, 0)
                    frmGGL.rihNr.SelStart = 0
                    frmGGL.rihNr.SelLength = 0
            
                
                If frmGGL.lblZZ.Caption = mod1.DName Or mod1.DName = "马晓聪" Then
                frmGGL.cmdDel.Enabled = True
                Else
                frmGGL.cmdDel.Enabled = False
                End If
                

                frmGGL.cmdYjb.Visible = False
                    timFl.Enabled = True
            
            End If
              
              
              
            
                    'modGGL.Oid = Gid
                    MDI.ztT.Panels(3).Text = zz & "刚给您发了一条短信，请注意收看！"
                    'MDI.ztT.Panels(3).Style = sbrCaps
                    'Call Fl(1000)

    Else

        
        'Call Fl(1)
        timFl.Enabled = False
  End If
End Sub



Public Sub Fl(FF As Integer)
      Dim FlashInfo     As FLASHWINFO
        If FF = 1 Then
            Sleep (10000)
            FlashWindowEx FlashInfo
        Else
            'Specifies   the   size   of   the   structure.
            FlashInfo.cbSize = Len(FlashInfo)
            'Specifies   the   flash   status
            FlashInfo.dwFlags = FLASHW_ALL Or FLASHW_TIMER
            'Specifies   the   rate,   in   milliseconds,   at   which   the   window   will   be   flashed.   If _
              dwTimeout   is   zero,   the   function   uses   the   default   cursor   blink   rate.
            FlashInfo.dwTimeout = 0
            'Handle   to   the   window   to   be   flashed.   The   window   can   be   either   opened   or   minimized.
            FlashInfo.hwnd = Me.hwnd
            'Specifies   the   number   of   times   to   flash   the   window.
            FlashInfo.uCount = FF
            Sleep (0)
            FlashWindowEx FlashInfo
      End If
        
End Sub
