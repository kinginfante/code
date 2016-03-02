VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form wbDN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "维保客户机密档案"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   15270
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "审核"
      Height          =   765
      Left            =   12510
      Picture         =   "wbDN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   298
      Top             =   8460
      Width           =   675
   End
   Begin VB.Frame frmQM 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   2250
      TabIndex        =   292
      Top             =   7350
      Visible         =   0   'False
      Width           =   6315
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1365
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   296
         Top             =   300
         Width           =   4965
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   295
         Top             =   480
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   294
         Top             =   870
         Width           =   675
      End
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5250
         Style           =   1  'Graphical
         TabIndex        =   293
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Timer timQuit 
      Interval        =   1000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer timWait 
      Interval        =   1000
      Left            =   60
      Top             =   30
   End
   Begin VB.Frame frmLblQT 
      Caption         =   "Frame1"
      Height          =   1305
      Left            =   12870
      TabIndex        =   279
      Top             =   30
      Width           =   2415
      Begin VB.OptionButton lblQT 
         Caption         =   "lblQT"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   281
         Top             =   210
         Width           =   2265
      End
      Begin VB.OptionButton lblQT 
         Caption         =   "lblQT"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   282
         Top             =   435
         Width           =   2265
      End
      Begin VB.OptionButton lblQT 
         Caption         =   "lblQT"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   280
         Top             =   1080
         Width           =   2265
      End
      Begin VB.OptionButton lblQT 
         Caption         =   "lblQT"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   283
         Top             =   645
         Width           =   2265
      End
      Begin VB.OptionButton lblQT 
         Caption         =   "lblQT"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   284
         Top             =   870
         Width           =   2265
      End
   End
   Begin VB.TextBox txtXmAdr 
      DataField       =   "xmAdr"
      DataSource      =   "adodm1"
      Height          =   285
      Left            =   1290
      TabIndex        =   253
      Top             =   960
      Width           =   4635
   End
   Begin VB.OptionButton optQt 
      Caption         =   "其它客户"
      Height          =   1005
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   246
      Top             =   180
      Width           =   315
   End
   Begin VB.OptionButton optWy 
      Caption         =   "物业"
      Height          =   255
      Left            =   9210
      Style           =   1  'Graphical
      TabIndex        =   245
      Top             =   570
      Width           =   765
   End
   Begin VB.OptionButton optYz 
      Caption         =   "业主"
      Height          =   285
      Left            =   9210
      Style           =   1  'Graphical
      TabIndex        =   244
      Top             =   150
      Width           =   765
   End
   Begin VB.TextBox txtXmmc 
      Height          =   315
      Left            =   1290
      TabIndex        =   237
      Top             =   90
      Width           =   4635
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      CausesValidation=   0   'False
      Height          =   765
      Left            =   13890
      Picture         =   "wbDN.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   234
      ToolTipText     =   "提交"
      Top             =   8460
      Width           =   675
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "退出"
      CausesValidation=   0   'False
      Height          =   765
      Left            =   14580
      Picture         =   "wbDN.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   233
      ToolTipText     =   "返回"
      Top             =   8460
      Width           =   675
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      CausesValidation=   0   'False
      Height          =   765
      Left            =   13200
      Picture         =   "wbDN.frx":0BAE
      Style           =   1  'Graphical
      TabIndex        =   232
      ToolTipText     =   "修改"
      Top             =   8460
      Width           =   675
   End
   Begin TabDlg.SSTab tabKh 
      Height          =   7875
      Left            =   -30
      TabIndex        =   0
      Top             =   1350
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   13891
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "客户单位"
      TabPicture(0)   =   "wbDN.frx":0FF0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtKhmc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "comXyxz"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "comQy"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtFH"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtKhYY"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtZH"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtKhDm"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "comXz"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtAdr1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "frmGL"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "frmJz"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblKid"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label21"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label14(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label6"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label4"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label8(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label12"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label15"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label13"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblgdate(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblgdate(4)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblgdate(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "联系人档案"
      TabPicture(1)   =   "wbDN.frx":100C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "tabRen"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "adoRen"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "dtgRen"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdRdel"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdRadd"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdRight"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdLeft"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdNew"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdQing"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.CommandButton cmdQing 
         Caption         =   "清空"
         Height          =   435
         Left            =   12960
         TabIndex        =   289
         Top             =   6780
         Width           =   555
      End
      Begin MSDataListLib.DataCombo txtKhmc 
         Height          =   330
         Left            =   -73680
         TabIndex        =   286
         Top             =   480
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   582
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "添加"
         Height          =   435
         Left            =   13560
         TabIndex        =   226
         Top             =   6780
         Width           =   555
      End
      Begin VB.TextBox comXyxz 
         Height          =   330
         Left            =   -68310
         Locked          =   -1  'True
         TabIndex        =   212
         Top             =   1410
         Width           =   3435
      End
      Begin VB.CommandButton cmdLeft 
         Caption         =   "<"
         Height          =   405
         Left            =   10770
         TabIndex        =   23
         Top             =   6840
         Width           =   555
      End
      Begin VB.CommandButton cmdRight 
         Caption         =   ">"
         Height          =   405
         Left            =   11550
         TabIndex        =   22
         Top             =   6810
         Width           =   555
      End
      Begin VB.CommandButton cmdRadd 
         Caption         =   "更新"
         Height          =   435
         Left            =   14130
         TabIndex        =   21
         Top             =   6780
         Width           =   555
      End
      Begin VB.CommandButton cmdRdel 
         Caption         =   "删除"
         Height          =   435
         Left            =   14670
         TabIndex        =   20
         Top             =   6780
         Width           =   555
      End
      Begin MSDataListLib.DataCombo comQy 
         Height          =   330
         Left            =   -68310
         TabIndex        =   19
         Top             =   480
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox txtFH 
         DataField       =   "FH"
         Height          =   285
         Left            =   -63540
         TabIndex        =   12
         Top             =   480
         Width           =   3555
      End
      Begin VB.TextBox txtKhYY 
         DataField       =   "khYY"
         Height          =   330
         Left            =   -68310
         TabIndex        =   11
         Top             =   960
         Width           =   3435
      End
      Begin VB.TextBox txtZH 
         DataField       =   "ZH"
         Height          =   285
         Left            =   -63540
         TabIndex        =   10
         Top             =   945
         Width           =   3585
      End
      Begin VB.TextBox txtKhDm 
         BackColor       =   &H00FFFFFF&
         DataField       =   "khDh"
         DataSource      =   "adodm1"
         Height          =   300
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "txtKhDm"
         Top             =   960
         Width           =   3795
      End
      Begin VB.ComboBox comXz 
         DataField       =   "qyXz"
         DataSource      =   "adodm1"
         Height          =   300
         ItemData        =   "wbDN.frx":1028
         Left            =   -73680
         List            =   "wbDN.frx":103B
         TabIndex        =   8
         Top             =   1410
         Width           =   3825
      End
      Begin VB.TextBox txtAdr1 
         DataField       =   "xmAdr"
         DataSource      =   "adodm1"
         Height          =   285
         Left            =   -63540
         TabIndex        =   5
         Top             =   1410
         Width           =   3585
      End
      Begin MSDataGridLib.DataGrid dtgRen 
         Bindings        =   "wbDN.frx":105D
         Height          =   6375
         Left            =   9990
         TabIndex        =   24
         Top             =   300
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   11245
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   18
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
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "联系人列表"
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "khMan"
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
         BeginProperty Column01 
            DataField       =   "khZw"
            Caption         =   "职务"
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
            DataField       =   "rid"
            Caption         =   "rid"
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
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column02 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc adoRen 
         Height          =   330
         Left            =   9060
         Top             =   6990
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
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
      Begin TabDlg.SSTab tabRen 
         Height          =   7665
         Left            =   0
         TabIndex        =   25
         Top             =   240
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   13520
         _Version        =   393216
         Tabs            =   6
         TabHeight       =   520
         TabCaption(0)   =   "客户"
         TabPicture(0)   =   "wbDN.frx":1072
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label5(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label14(5)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label10(2)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label14(3)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label33(0)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblgdate(15)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label24"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label26(0)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label35(0)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label37(0)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label14(4)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label19"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label10(0)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "Label14(6)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label8(0)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "lblLhk(0)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label20"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "Label23"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "Label25"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "lblXb"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "lblQM(0)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "lblTm(0)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "lblLcUid"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "lblLcRen"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "lblLc"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "lblRid"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "lblUid"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "lblYwy"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "lblXuid"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "lblXywy"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "lblFwid"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "Label14(2)"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "lblgdate(10)"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "Label16"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "Label18"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "txtK(4)"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "dtpSr"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "txtMan"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "txtK(0)"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).Control(40)=   "txtZw"
         Tab(0).Control(40).Enabled=   0   'False
         Tab(0).Control(41)=   "txtLjadr"
         Tab(0).Control(41).Enabled=   0   'False
         Tab(0).Control(42)=   "txtLpho"
         Tab(0).Control(42).Enabled=   0   'False
         Tab(0).Control(43)=   "txtLjpho"
         Tab(0).Control(43).Enabled=   0   'False
         Tab(0).Control(44)=   "txtLdwdz"
         Tab(0).Control(44).Enabled=   0   'False
         Tab(0).Control(45)=   "txtLjmob"
         Tab(0).Control(45).Enabled=   0   'False
         Tab(0).Control(46)=   "txtK(1)"
         Tab(0).Control(46).Enabled=   0   'False
         Tab(0).Control(47)=   "optMan"
         Tab(0).Control(47).Enabled=   0   'False
         Tab(0).Control(48)=   "optWoman"
         Tab(0).Control(48).Enabled=   0   'False
         Tab(0).Control(49)=   "txtHk"
         Tab(0).Control(49).Enabled=   0   'False
         Tab(0).Control(50)=   "txtK(2)"
         Tab(0).Control(50).Enabled=   0   'False
         Tab(0).Control(51)=   "txtK(3)"
         Tab(0).Control(51).Enabled=   0   'False
         Tab(0).Control(52)=   "txtSr"
         Tab(0).Control(52).Enabled=   0   'False
         Tab(0).Control(53)=   "cmdQm(0)"
         Tab(0).Control(53).Enabled=   0   'False
         Tab(0).Control(54)=   "txtK(80)"
         Tab(0).Control(54).Enabled=   0   'False
         Tab(0).Control(55)=   "txtK(81)"
         Tab(0).Control(55).Enabled=   0   'False
         Tab(0).Control(56)=   "txtK(82)"
         Tab(0).Control(56).Enabled=   0   'False
         Tab(0).ControlCount=   57
         TabCaption(1)   =   "教育背景"
         TabPicture(1)   =   "wbDN.frx":108E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtK(7)"
         Tab(1).Control(1)=   "txtK(6)"
         Tab(1).Control(2)=   "txtK(12)"
         Tab(1).Control(3)=   "txtK(11)"
         Tab(1).Control(4)=   "txtK(10)"
         Tab(1).Control(5)=   "txtK(9)"
         Tab(1).Control(6)=   "txtK(8)"
         Tab(1).Control(7)=   "dtpBy"
         Tab(1).Control(8)=   "txtK(5)"
         Tab(1).Control(9)=   "Label47"
         Tab(1).Control(10)=   "Label46"
         Tab(1).Control(11)=   "Label45"
         Tab(1).Control(12)=   "Label44"
         Tab(1).Control(13)=   "Label43"
         Tab(1).Control(14)=   "Label42"
         Tab(1).Control(15)=   "Label41"
         Tab(1).Control(16)=   "Label40"
         Tab(1).ControlCount=   17
         TabCaption(2)   =   "家庭"
         TabPicture(2)   =   "wbDN.frx":10AA
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtK(17)"
         Tab(2).Control(1)=   "txtK(21)"
         Tab(2).Control(2)=   "txtK(20)"
         Tab(2).Control(3)=   "txtK(19)"
         Tab(2).Control(4)=   "dtpJh"
         Tab(2).Control(5)=   "txtK(16)"
         Tab(2).Control(6)=   "txtK(15)"
         Tab(2).Control(7)=   "txtK(14)"
         Tab(2).Control(8)=   "txtK(13)"
         Tab(2).Control(9)=   "txtK(18)"
         Tab(2).Control(10)=   "Label56"
         Tab(2).Control(11)=   "Label55"
         Tab(2).Control(12)=   "Label54"
         Tab(2).Control(13)=   "Label53"
         Tab(2).Control(14)=   "Label52"
         Tab(2).Control(15)=   "Label51"
         Tab(2).Control(16)=   "Label50"
         Tab(2).Control(17)=   "Label49"
         Tab(2).Control(18)=   "Label48"
         Tab(2).ControlCount=   19
         TabCaption(3)   =   "业务背景资料"
         TabPicture(3)   =   "wbDN.frx":10C6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtK(24)"
         Tab(3).Control(1)=   "txtK(27)"
         Tab(3).Control(2)=   "txtK(37)"
         Tab(3).Control(3)=   "txtK(36)"
         Tab(3).Control(4)=   "txtK(35)"
         Tab(3).Control(5)=   "txtK(34)"
         Tab(3).Control(6)=   "txtK(33)"
         Tab(3).Control(7)=   "txtK(32)"
         Tab(3).Control(8)=   "txtK(31)"
         Tab(3).Control(9)=   "txtK(30)"
         Tab(3).Control(10)=   "txtK(29)"
         Tab(3).Control(11)=   "txtK(28)"
         Tab(3).Control(12)=   "txtK(26)"
         Tab(3).Control(13)=   "txtK(25)"
         Tab(3).Control(14)=   "txtK(23)"
         Tab(3).Control(15)=   "txtK(22)"
         Tab(3).Control(16)=   "Label84"
         Tab(3).Control(17)=   "Label83"
         Tab(3).Control(18)=   "Label82"
         Tab(3).Control(19)=   "Label71"
         Tab(3).Control(20)=   "Label70"
         Tab(3).Control(21)=   "Label69"
         Tab(3).Control(22)=   "Label68"
         Tab(3).Control(23)=   "Label67"
         Tab(3).Control(24)=   "Label66"
         Tab(3).Control(25)=   "Label65"
         Tab(3).Control(26)=   "Line2"
         Tab(3).Control(27)=   "Label64"
         Tab(3).Control(28)=   "Label63"
         Tab(3).Control(29)=   "Label62"
         Tab(3).Control(30)=   "Line1"
         Tab(3).Control(31)=   "Label61"
         Tab(3).Control(32)=   "Label60"
         Tab(3).Control(33)=   "Label59"
         Tab(3).Control(34)=   "Label58"
         Tab(3).Control(35)=   "Label57"
         Tab(3).ControlCount=   36
         TabCaption(4)   =   "生活方式"
         TabPicture(4)   =   "wbDN.frx":10E2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "txtK(79)"
         Tab(4).Control(1)=   "txtK(65)"
         Tab(4).Control(2)=   "txtK(64)"
         Tab(4).Control(3)=   "txtK(63)"
         Tab(4).Control(4)=   "txtK(61)"
         Tab(4).Control(5)=   "txtK(62)"
         Tab(4).Control(6)=   "txtK(60)"
         Tab(4).Control(7)=   "txtK(59)"
         Tab(4).Control(8)=   "txtK(58)"
         Tab(4).Control(9)=   "txtK(57)"
         Tab(4).Control(10)=   "txtK(56)"
         Tab(4).Control(11)=   "Text60"
         Tab(4).Control(12)=   "txtK(55)"
         Tab(4).Control(13)=   "txtK(53)"
         Tab(4).Control(14)=   "txtK(51)"
         Tab(4).Control(15)=   "txtK(49)"
         Tab(4).Control(16)=   "txtK(54)"
         Tab(4).Control(17)=   "txtK(52)"
         Tab(4).Control(18)=   "Text53"
         Tab(4).Control(19)=   "Text51"
         Tab(4).Control(20)=   "txtK(50)"
         Tab(4).Control(21)=   "txtK(48)"
         Tab(4).Control(22)=   "txtK(47)"
         Tab(4).Control(23)=   "txtK(46)"
         Tab(4).Control(24)=   "txtK(45)"
         Tab(4).Control(25)=   "txtK(44)"
         Tab(4).Control(26)=   "txtK(43)"
         Tab(4).Control(27)=   "txtK(42)"
         Tab(4).Control(28)=   "txtK(41)"
         Tab(4).Control(29)=   "txtK(40)"
         Tab(4).Control(30)=   "txtK(39)"
         Tab(4).Control(31)=   "txtK(38)"
         Tab(4).Control(32)=   "Text52"
         Tab(4).Control(33)=   "lblgdate(16)"
         Tab(4).Control(34)=   "Label102"
         Tab(4).Control(35)=   "Label100"
         Tab(4).Control(36)=   "Label97"
         Tab(4).Control(37)=   "Line5"
         Tab(4).Control(38)=   "Line4"
         Tab(4).Control(39)=   "Line3"
         Tab(4).Control(40)=   "Label81"
         Tab(4).Control(41)=   "Label80"
         Tab(4).Control(42)=   "Label79"
         Tab(4).Control(43)=   "Label78"
         Tab(4).Control(44)=   "Label77"
         Tab(4).Control(45)=   "Label76"
         Tab(4).Control(46)=   "Label75"
         Tab(4).Control(47)=   "Label74"
         Tab(4).Control(48)=   "Label73"
         Tab(4).Control(49)=   "Label72"
         Tab(4).ControlCount=   50
         TabCaption(5)   =   "客户与你"
         TabPicture(5)   =   "wbDN.frx":10FE
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Label9"
         Tab(5).Control(1)=   "Label11"
         Tab(5).Control(2)=   "Text71"
         Tab(5).Control(3)=   "txtK(68)"
         Tab(5).Control(4)=   "txtK(70)"
         Tab(5).Control(5)=   "txtK(71)"
         Tab(5).Control(6)=   "Text75"
         Tab(5).Control(7)=   "txtK(73)"
         Tab(5).Control(8)=   "txtK(74)"
         Tab(5).Control(9)=   "txtK(75)"
         Tab(5).Control(10)=   "txtK(76)"
         Tab(5).Control(11)=   "txtK(77)"
         Tab(5).Control(12)=   "txtK(78)"
         Tab(5).Control(13)=   "txtK(72)"
         Tab(5).Control(14)=   "txtK(69)"
         Tab(5).Control(15)=   "txtK(67)"
         Tab(5).Control(16)=   "txtK(66)"
         Tab(5).ControlCount=   17
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   82
            Left            =   3990
            TabIndex        =   256
            Tag             =   "6"
            Text            =   "82"
            Top             =   3300
            Width           =   2025
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   79
            Left            =   -72750
            TabIndex        =   235
            Tag             =   "50"
            Text            =   "79"
            Top             =   3780
            Width           =   7065
         End
         Begin VB.TextBox txtK 
            DataField       =   "khYb1"
            DataSource      =   "adodm1"
            Height          =   285
            Index           =   81
            Left            =   3990
            TabIndex        =   228
            Tag             =   "6"
            Text            =   "81"
            Top             =   2640
            Width           =   1995
         End
         Begin VB.TextBox txtK 
            DataField       =   "khCz1"
            DataSource      =   "adodm1"
            Height          =   315
            Index           =   80
            Left            =   7410
            TabIndex        =   227
            Tag             =   "15"
            Text            =   "80"
            Top             =   2640
            Width           =   1995
         End
         Begin VB.CommandButton cmdQm 
            Caption         =   "cmdQm"
            Height          =   465
            Index           =   0
            Left            =   1050
            TabIndex        =   213
            Top             =   6360
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.TextBox txtSr 
            Height          =   300
            Left            =   1110
            Locked          =   -1  'True
            TabIndex        =   196
            Top             =   3840
            Width           =   1455
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   24
            Left            =   -72840
            TabIndex        =   195
            Tag             =   "20"
            Text            =   "24"
            Top             =   1650
            Width           =   2775
         End
         Begin VB.TextBox txtK 
            Height          =   345
            Index           =   27
            Left            =   -69960
            TabIndex        =   194
            Tag             =   "20"
            Text            =   "27"
            Top             =   2370
            Width           =   4365
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   17
            Left            =   -72690
            Locked          =   -1  'True
            TabIndex        =   193
            Tag             =   "20"
            Text            =   "17"
            Top             =   3810
            Width           =   6735
         End
         Begin VB.TextBox txtK 
            Height          =   330
            Index           =   7
            Left            =   -72840
            Locked          =   -1  'True
            TabIndex        =   192
            Text            =   "7"
            Top             =   1980
            Width           =   2085
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   66
            Left            =   -71700
            TabIndex        =   191
            Tag             =   "30"
            Text            =   "66"
            Top             =   750
            Width           =   5985
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   67
            Left            =   -71700
            TabIndex        =   190
            Tag             =   "30"
            Text            =   "67"
            Top             =   1260
            Width           =   5985
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   69
            Left            =   -71700
            TabIndex        =   189
            Tag             =   "30"
            Text            =   "69"
            Top             =   2340
            Width           =   5985
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   72
            Left            =   -71700
            TabIndex        =   188
            Tag             =   "30"
            Text            =   "72"
            Top             =   3450
            Width           =   5985
         End
         Begin VB.TextBox txtK 
            Height          =   585
            Index           =   78
            Left            =   -73350
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   187
            Tag             =   "100"
            Text            =   "wbDN.frx":111A
            Top             =   6750
            Width           =   7635
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   77
            Left            =   -70830
            TabIndex        =   186
            Tag             =   "30"
            Text            =   "77"
            Top             =   6180
            Width           =   5115
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   76
            Left            =   -71700
            TabIndex        =   185
            Tag             =   "30"
            Text            =   "76"
            Top             =   5640
            Width           =   5985
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   75
            Left            =   -71700
            TabIndex        =   184
            Tag             =   "30"
            Text            =   "75"
            Top             =   5085
            Width           =   5985
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   74
            Left            =   -71700
            TabIndex        =   183
            Tag             =   "30"
            Text            =   "74"
            Top             =   4560
            Width           =   5985
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   73
            Left            =   -71700
            TabIndex        =   182
            Tag             =   "30"
            Text            =   "73"
            Top             =   3975
            Width           =   5985
         End
         Begin VB.TextBox Text75 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   3855
            Left            =   -74730
            MultiLine       =   -1  'True
            TabIndex        =   181
            Text            =   "wbDN.frx":111D
            Top             =   3540
            Width           =   4065
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   71
            Left            =   -68820
            TabIndex        =   180
            Tag             =   "20"
            Text            =   "71"
            Top             =   2880
            Width           =   3135
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   70
            Left            =   -72780
            TabIndex        =   178
            Tag             =   "20"
            Text            =   "70"
            Top             =   2910
            Width           =   1665
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   68
            Left            =   -70620
            TabIndex        =   176
            Tag             =   "30"
            Text            =   "68"
            Top             =   1830
            Width           =   4935
         End
         Begin VB.TextBox Text71 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   1965
            Left            =   -74730
            MultiLine       =   -1  'True
            TabIndex        =   175
            Text            =   "wbDN.frx":1201
            Top             =   810
            Width           =   3975
         End
         Begin VB.TextBox txtK 
            Height          =   270
            Index           =   65
            Left            =   -70620
            TabIndex        =   174
            Tag             =   "30"
            Text            =   "65"
            Top             =   6810
            Width           =   4965
         End
         Begin VB.TextBox txtK 
            Height          =   270
            Index           =   64
            Left            =   -70620
            TabIndex        =   173
            Tag             =   "30"
            Text            =   "64"
            Top             =   6450
            Width           =   4965
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   63
            Left            =   -68130
            TabIndex        =   172
            Tag             =   "20"
            Text            =   "63"
            Top             =   6090
            Width           =   2475
         End
         Begin VB.TextBox txtK 
            Height          =   285
            Index           =   61
            Left            =   -68130
            TabIndex        =   171
            Tag             =   "30"
            Text            =   "61"
            Top             =   5760
            Width           =   2475
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   62
            Left            =   -72690
            TabIndex        =   170
            Tag             =   "20"
            Text            =   "62"
            Top             =   6090
            Width           =   3375
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   60
            Left            =   -72690
            TabIndex        =   169
            Tag             =   "20"
            Text            =   "60"
            Top             =   5730
            Width           =   3375
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   59
            Left            =   -72690
            TabIndex        =   168
            Tag             =   "30"
            Text            =   "59"
            Top             =   5370
            Width           =   7035
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   58
            Left            =   -71940
            TabIndex        =   167
            Tag             =   "30"
            Text            =   "58"
            Top             =   4950
            Width           =   6285
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   57
            Left            =   -71940
            TabIndex        =   166
            Tag             =   "30"
            Text            =   "57"
            Top             =   4590
            Width           =   6285
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   56
            Left            =   -71940
            TabIndex        =   165
            Tag             =   "30"
            Text            =   "56"
            Top             =   4200
            Width           =   6285
         End
         Begin VB.TextBox Text60 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   1665
            Left            =   -74610
            MultiLine       =   -1  'True
            TabIndex        =   164
            Text            =   "wbDN.frx":128A
            Top             =   5430
            Width           =   3885
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   55
            Left            =   -68070
            TabIndex        =   163
            Tag             =   "20"
            Text            =   "55"
            Top             =   3450
            Width           =   2385
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   53
            Left            =   -68070
            TabIndex        =   162
            Tag             =   "20"
            Text            =   "53"
            Top             =   3090
            Width           =   2385
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   51
            Left            =   -68070
            TabIndex        =   161
            Tag             =   "20"
            Text            =   "51"
            Top             =   2730
            Width           =   2385
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   49
            Left            =   -68070
            TabIndex        =   160
            Tag             =   "20"
            Text            =   "49"
            Top             =   2340
            Width           =   2385
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   54
            Left            =   -72750
            TabIndex        =   159
            Tag             =   "20"
            Text            =   "54"
            Top             =   3450
            Width           =   2385
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   52
            Left            =   -72750
            TabIndex        =   158
            Tag             =   "20"
            Text            =   "52"
            Top             =   3075
            Width           =   2385
         End
         Begin VB.TextBox Text53 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   -70140
            MultiLine       =   -1  'True
            TabIndex        =   157
            Text            =   "wbDN.frx":1327
            Top             =   2400
            Width           =   1965
         End
         Begin VB.TextBox Text51 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   945
            Left            =   -74640
            MultiLine       =   -1  'True
            TabIndex        =   154
            Text            =   "wbDN.frx":1386
            Top             =   4260
            Width           =   2505
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   50
            Left            =   -72750
            TabIndex        =   153
            Tag             =   "20"
            Text            =   "50"
            Top             =   2715
            Width           =   2385
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   48
            Left            =   -72750
            TabIndex        =   152
            Tag             =   "20"
            Text            =   "48"
            Top             =   2340
            Width           =   2385
         End
         Begin VB.TextBox txtK 
            Height          =   495
            Index           =   37
            Left            =   -68430
            TabIndex        =   148
            Tag             =   "30"
            Text            =   "37"
            Top             =   6510
            Width           =   2805
         End
         Begin VB.TextBox txtK 
            Height          =   465
            Index           =   36
            Left            =   -71760
            TabIndex        =   146
            Tag             =   "20"
            Text            =   "36"
            Top             =   6510
            Width           =   2265
         End
         Begin VB.TextBox txtK 
            Height          =   405
            Index           =   35
            Left            =   -71760
            TabIndex        =   144
            Tag             =   "30"
            Text            =   "35"
            Top             =   6090
            Width           =   6105
         End
         Begin VB.TextBox txtK 
            Height          =   270
            Index           =   47
            Left            =   -68190
            TabIndex        =   142
            Tag             =   "20"
            Text            =   "47"
            Top             =   1920
            Width           =   2565
         End
         Begin VB.TextBox txtK 
            Height          =   270
            Index           =   46
            Left            =   -73200
            TabIndex        =   140
            Tag             =   "20"
            Text            =   "46"
            Top             =   1980
            Width           =   2385
         End
         Begin VB.TextBox txtK 
            Height          =   285
            Index           =   45
            Left            =   -68190
            TabIndex        =   138
            Tag             =   "20"
            Text            =   "45"
            Top             =   1620
            Width           =   2565
         End
         Begin VB.TextBox txtK 
            Height          =   270
            Index           =   44
            Left            =   -73200
            TabIndex        =   136
            Tag             =   "20"
            Text            =   "44"
            Top             =   1680
            Width           =   2385
         End
         Begin VB.TextBox txtK 
            Height          =   285
            Index           =   43
            Left            =   -69540
            TabIndex        =   134
            Tag             =   "20"
            Text            =   "43"
            Top             =   1320
            Width           =   3885
         End
         Begin VB.TextBox txtK 
            Height          =   285
            Index           =   42
            Left            =   -73860
            TabIndex        =   132
            Tag             =   "10"
            Text            =   "42"
            Top             =   1350
            Width           =   2055
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   41
            Left            =   -67230
            TabIndex        =   130
            Tag             =   "20"
            Text            =   "41"
            Top             =   990
            Width           =   1575
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   40
            Left            =   -71160
            TabIndex        =   128
            Tag             =   "20"
            Text            =   "40"
            Top             =   1020
            Width           =   1215
         End
         Begin VB.TextBox txtK 
            Height          =   285
            Index           =   39
            Left            =   -73860
            TabIndex        =   126
            Tag             =   "10"
            Text            =   "39"
            Top             =   1020
            Width           =   1155
         End
         Begin VB.TextBox txtK 
            Height          =   285
            Index           =   38
            Left            =   -72750
            TabIndex        =   124
            Tag             =   "30"
            Text            =   "38"
            Top             =   660
            Width           =   7095
         End
         Begin VB.TextBox txtK 
            Height          =   465
            Index           =   34
            Left            =   -71760
            TabIndex        =   122
            Tag             =   "30"
            Text            =   "34"
            Top             =   5610
            Width           =   6135
         End
         Begin VB.TextBox txtK 
            Height          =   465
            Index           =   33
            Left            =   -71760
            TabIndex        =   121
            Tag             =   "30"
            Text            =   "33"
            Top             =   5100
            Width           =   6135
         End
         Begin VB.TextBox txtK 
            Height          =   405
            Index           =   32
            Left            =   -71760
            TabIndex        =   120
            Tag             =   "30"
            Text            =   "32"
            Top             =   4650
            Width           =   6135
         End
         Begin VB.TextBox txtK 
            Height          =   555
            Index           =   31
            Left            =   -71760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   116
            Tag             =   "30"
            Text            =   "wbDN.frx":13D9
            Top             =   4050
            Width           =   6135
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   30
            Left            =   -71760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   115
            Tag             =   "30"
            Text            =   "wbDN.frx":13DE
            Top             =   3660
            Width           =   6135
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   29
            Left            =   -71760
            TabIndex        =   114
            Tag             =   "20"
            Text            =   "29"
            Top             =   3240
            Width           =   6135
         End
         Begin VB.TextBox txtK 
            Height          =   345
            Index           =   28
            Left            =   -71760
            TabIndex        =   113
            Tag             =   "20"
            Text            =   "28"
            Top             =   2850
            Width           =   6135
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   26
            Left            =   -72840
            TabIndex        =   112
            Tag             =   "10"
            Text            =   "26"
            Top             =   2370
            Width           =   1695
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   25
            Left            =   -69030
            TabIndex        =   111
            Tag             =   "10"
            Text            =   "25"
            Top             =   1650
            Width           =   3405
         End
         Begin VB.TextBox txtK 
            Height          =   345
            Index           =   23
            Left            =   -72840
            TabIndex        =   110
            Tag             =   "30"
            Text            =   "23"
            Top             =   1260
            Width           =   7215
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   22
            Left            =   -72840
            TabIndex        =   109
            Tag             =   "20"
            Text            =   "22"
            Top             =   870
            Width           =   7215
         End
         Begin VB.TextBox txtK 
            Height          =   555
            Index           =   21
            Left            =   -72690
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   96
            Tag             =   "50"
            Text            =   "wbDN.frx":13E3
            Top             =   5880
            Width           =   7035
         End
         Begin VB.TextBox txtK 
            Height          =   525
            Index           =   20
            Left            =   -72690
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   95
            Tag             =   "50"
            Text            =   "wbDN.frx":13E6
            Top             =   5190
            Width           =   7035
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   19
            Left            =   -68370
            TabIndex        =   94
            Tag             =   "5"
            Text            =   "19"
            Top             =   4560
            Width           =   2715
         End
         Begin MSComCtl2.DTPicker dtpJh 
            Height          =   375
            Left            =   -72690
            TabIndex        =   93
            Top             =   3810
            Width           =   7065
            _ExtentX        =   12462
            _ExtentY        =   661
            _Version        =   393216
            Format          =   133365761
            CurrentDate     =   38712
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   16
            Left            =   -72690
            TabIndex        =   92
            Tag             =   "30"
            Text            =   "16"
            Top             =   2970
            Width           =   7035
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   15
            Left            =   -72690
            TabIndex        =   91
            Tag             =   "10"
            Text            =   "15"
            Top             =   2250
            Width           =   7035
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   14
            Left            =   -72690
            TabIndex        =   90
            Tag             =   "10"
            Text            =   "14"
            Top             =   1560
            Width           =   7035
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   13
            Left            =   -72690
            TabIndex        =   89
            Tag             =   "30"
            Text            =   "13"
            Top             =   840
            Width           =   7035
         End
         Begin VB.TextBox txtK 
            Height          =   375
            Index           =   18
            Left            =   -72690
            TabIndex        =   85
            Tag             =   "5"
            Text            =   "18"
            Top             =   4590
            Width           =   2355
         End
         Begin VB.TextBox txtK 
            Height          =   405
            Index           =   6
            Left            =   -72840
            TabIndex        =   78
            Tag             =   "20"
            Text            =   "6"
            Top             =   1410
            Width           =   7155
         End
         Begin VB.TextBox txtK 
            Height          =   1065
            Index           =   12
            Left            =   -72840
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   77
            Tag             =   "50"
            Text            =   "wbDN.frx":13E9
            Top             =   4890
            Width           =   7125
         End
         Begin VB.TextBox txtK 
            Height          =   675
            Index           =   11
            Left            =   -72840
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   75
            Tag             =   "30"
            Text            =   "wbDN.frx":13EE
            Top             =   3870
            Width           =   7125
         End
         Begin VB.TextBox txtK 
            Height          =   345
            Index           =   10
            Left            =   -72840
            TabIndex        =   73
            Tag             =   "20"
            Text            =   "10"
            Top             =   3300
            Width           =   7095
         End
         Begin VB.TextBox txtK 
            Height          =   345
            Index           =   9
            Left            =   -72840
            TabIndex        =   71
            Tag             =   "30"
            Text            =   "9"
            Top             =   2610
            Width           =   7125
         End
         Begin VB.TextBox txtK 
            Height          =   345
            Index           =   8
            Left            =   -68670
            TabIndex        =   69
            Tag             =   "10"
            Text            =   "8"
            Top             =   1950
            Width           =   2985
         End
         Begin MSComCtl2.DTPicker dtpBy 
            Height          =   345
            Left            =   -72840
            TabIndex        =   68
            Top             =   1980
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   609
            _Version        =   393216
            Format          =   133365761
            CurrentDate     =   38712
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   -360
            TabIndex        =   65
            Text            =   "Text10"
            Top             =   -2970
            Width           =   7155
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   5
            Left            =   -72840
            TabIndex        =   64
            Tag             =   "30"
            Text            =   "5"
            Top             =   900
            Width           =   7155
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   3
            Left            =   3900
            TabIndex        =   61
            Tag             =   "5"
            Text            =   "3"
            Top             =   4500
            Width           =   1725
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   2
            Left            =   1110
            TabIndex        =   60
            Tag             =   "5"
            Text            =   "2"
            Top             =   4500
            Width           =   1695
         End
         Begin VB.TextBox txtHk 
            DataField       =   "khHk"
            Height          =   285
            Left            =   7410
            TabIndex        =   54
            Tag             =   "10"
            Top             =   4470
            Width           =   1995
         End
         Begin VB.OptionButton optWoman 
            Caption         =   "女"
            Height          =   285
            Left            =   4920
            TabIndex        =   51
            Top             =   3870
            Width           =   495
         End
         Begin VB.OptionButton optMan 
            Caption         =   "男"
            Height          =   195
            Left            =   4140
            TabIndex        =   50
            Top             =   3930
            Width           =   465
         End
         Begin VB.TextBox txtK 
            Height          =   315
            Index           =   1
            Left            =   7410
            TabIndex        =   47
            Tag             =   "20"
            Text            =   "1"
            Top             =   3840
            Width           =   1965
         End
         Begin VB.TextBox txtLjmob 
            DataField       =   "khMob"
            Height          =   315
            Left            =   7410
            TabIndex        =   43
            Tag             =   "11"
            Top             =   3300
            Width           =   2025
         End
         Begin VB.TextBox txtLdwdz 
            DataField       =   "khDwadr"
            Height          =   345
            Left            =   1110
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Tag             =   "50"
            Top             =   1410
            Width           =   8625
         End
         Begin VB.TextBox txtLjpho 
            DataField       =   "khJpho"
            Height          =   285
            Left            =   1110
            TabIndex        =   39
            Tag             =   "20"
            Top             =   3300
            Width           =   1695
         End
         Begin VB.TextBox txtLpho 
            DataField       =   "khDpho"
            Height          =   270
            Left            =   1110
            TabIndex        =   36
            Tag             =   "20"
            Top             =   2640
            Width           =   1785
         End
         Begin VB.TextBox txtLjadr 
            DataField       =   "khJadr"
            Height          =   330
            Left            =   1110
            TabIndex        =   34
            Tag             =   "50"
            Top             =   1980
            Width           =   8655
         End
         Begin VB.TextBox txtZw 
            DataField       =   "khZw"
            Height          =   315
            Left            =   6960
            TabIndex        =   31
            Tag             =   "20"
            Top             =   810
            Width           =   2745
         End
         Begin VB.TextBox txtK 
            Height          =   345
            Index           =   0
            Left            =   3900
            TabIndex        =   30
            Tag             =   "20"
            Text            =   "0"
            Top             =   780
            Width           =   1995
         End
         Begin VB.TextBox txtMan 
            DataField       =   "khMan"
            Height          =   285
            Left            =   1110
            TabIndex        =   26
            Tag             =   "10"
            Top             =   780
            Width           =   1785
         End
         Begin MSComCtl2.DTPicker dtpSr 
            Height          =   285
            Left            =   1110
            TabIndex        =   48
            Top             =   3840
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   503
            _Version        =   393216
            Format          =   133365761
            CurrentDate     =   38223
         End
         Begin VB.TextBox txtK 
            Height          =   555
            Index           =   4
            Left            =   1110
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   59
            Tag             =   "100"
            Text            =   "wbDN.frx":13F3
            Top             =   5250
            Width           =   8565
         End
         Begin VB.TextBox Text52 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   1695
            Left            =   -74880
            MultiLine       =   -1  'True
            TabIndex        =   156
            Text            =   "wbDN.frx":13F7
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label Label18 
            Caption         =   "邮编(宅)"
            Height          =   285
            Left            =   3120
            TabIndex        =   255
            Top             =   3330
            Width           =   765
         End
         Begin VB.Label lblgdate 
            Caption         =   "嗜好与娱乐"
            ForeColor       =   &H00C00000&
            Height          =   1155
            Index           =   16
            Left            =   -65580
            TabIndex        =   236
            Top             =   2460
            Width           =   345
         End
         Begin VB.Label Label16 
            Caption         =   "邮编(公)"
            Height          =   255
            Left            =   3120
            TabIndex        =   231
            Top             =   2670
            Width           =   735
         End
         Begin VB.Label lblgdate 
            Caption         =   "传    真"
            Height          =   255
            Index           =   10
            Left            =   6240
            TabIndex        =   230
            Top             =   2700
            Width           =   795
         End
         Begin VB.Label Label14 
            Caption         =   "*"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   2
            Left            =   7080
            TabIndex        =   229
            Top             =   2670
            Width           =   165
         End
         Begin VB.Label lblFwid 
            Caption         =   "lblFwid"
            Height          =   255
            Left            =   7710
            TabIndex        =   225
            Top             =   6750
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lblXywy 
            Caption         =   "lblXywy"
            Height          =   255
            Left            =   8880
            TabIndex        =   224
            Top             =   6630
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblXuid 
            Caption         =   "lblXuid"
            Height          =   255
            Left            =   8910
            TabIndex        =   223
            Top             =   6360
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lblYwy 
            Caption         =   "lblYwy"
            Height          =   225
            Left            =   7920
            TabIndex        =   222
            Top             =   6330
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lblUid 
            Caption         =   "lblUid"
            Height          =   285
            Left            =   7980
            TabIndex        =   221
            Top             =   6600
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lblRid 
            Caption         =   "lblRid"
            Height          =   225
            Left            =   5670
            TabIndex        =   220
            Top             =   6450
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label lblLc 
            Caption         =   "lblLc"
            Height          =   315
            Left            =   9000
            TabIndex        =   219
            Top             =   6120
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label lblLcRen 
            Caption         =   "lblLcRen"
            Height          =   285
            Left            =   6630
            TabIndex        =   218
            Top             =   6330
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lblLcUid 
            Caption         =   "lblLcUid"
            Height          =   285
            Left            =   6720
            TabIndex        =   217
            Top             =   6660
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label lblTm 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   0
            Left            =   990
            TabIndex        =   216
            Top             =   6870
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label lblQM 
            Caption         =   "lblQm"
            Height          =   225
            Index           =   0
            Left            =   1110
            TabIndex        =   214
            Top             =   6090
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label11 
            Caption         =   "是否道德感很强"
            Height          =   285
            Left            =   -70620
            TabIndex        =   179
            Top             =   2940
            Width           =   1635
         End
         Begin VB.Label Label9 
            Caption         =   "或以自我为中心"
            Height          =   345
            Left            =   -74700
            TabIndex        =   177
            Top             =   2940
            Width           =   1575
         End
         Begin VB.Label lblXb 
            Caption         =   "性别"
            DataField       =   "khSex"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "m/d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Left            =   4740
            TabIndex        =   155
            Top             =   4140
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label102 
            Caption         =   "是否热衷"
            Height          =   195
            Left            =   -69120
            TabIndex        =   151
            Top             =   6180
            Width           =   975
         End
         Begin VB.Label Label100 
            Caption         =   "如何参与"
            Height          =   255
            Left            =   -69120
            TabIndex        =   150
            Top             =   5820
            Width           =   975
         End
         Begin VB.Label Label97 
            Caption         =   "特殊兴趣"
            ForeColor       =   &H00FF0000&
            Height          =   945
            Left            =   -65580
            TabIndex        =   149
            Top             =   5400
            Width           =   225
         End
         Begin VB.Line Line5 
            X1              =   -74970
            X2              =   -65190
            Y1              =   5310
            Y2              =   5310
         End
         Begin VB.Line Line4 
            X1              =   -74970
            X2              =   -65190
            Y1              =   4170
            Y2              =   4170
         End
         Begin VB.Line Line3 
            X1              =   -75000
            X2              =   -65190
            Y1              =   2250
            Y2              =   2250
         End
         Begin VB.Label Label84 
            Caption         =   "为什么"
            Height          =   225
            Left            =   -69120
            TabIndex        =   147
            Top             =   6660
            Width           =   1095
         End
         Begin VB.Label Label83 
            Caption         =   "客户多思考现在或将来"
            Height          =   285
            Left            =   -74580
            TabIndex        =   145
            Top             =   6630
            Width           =   2235
         End
         Begin VB.Label Label82 
            Caption         =   "客户目前最关切的是公司前途或个人前途"
            Height          =   435
            Left            =   -74580
            TabIndex        =   143
            Top             =   6150
            Width           =   2415
         End
         Begin VB.Label Label81 
            Caption         =   "是否反对别人请客"
            Height          =   255
            Left            =   -70350
            TabIndex        =   141
            Top             =   1980
            Width           =   1665
         End
         Begin VB.Label Label80 
            Caption         =   "最偏好的菜式"
            Height          =   285
            Left            =   -74730
            TabIndex        =   139
            Top             =   1980
            Width           =   1335
         End
         Begin VB.Label Label79 
            Caption         =   "晚餐地点"
            Height          =   255
            Left            =   -69330
            TabIndex        =   137
            Top             =   1680
            Width           =   885
         End
         Begin VB.Label Label78 
            Caption         =   "最偏好的午餐地点"
            Height          =   255
            Left            =   -74760
            TabIndex        =   135
            Top             =   1710
            Width           =   1575
         End
         Begin VB.Label Label77 
            Caption         =   "若否，是否反对别人吸烟"
            Height          =   345
            Left            =   -71640
            TabIndex        =   133
            Top             =   1380
            Width           =   2175
         End
         Begin VB.Label Label76 
            Caption         =   "是否吸烟"
            Height          =   285
            Left            =   -74760
            TabIndex        =   131
            Top             =   1395
            Width           =   1095
         End
         Begin VB.Label Label75 
            Caption         =   "如果不喝酒，是否反对别人喝酒"
            Height          =   315
            Left            =   -69840
            TabIndex        =   129
            Top             =   1050
            Width           =   2625
         End
         Begin VB.Label Label74 
            Caption         =   "所嗜酒类与分量"
            Height          =   255
            Left            =   -72540
            TabIndex        =   127
            Top             =   1080
            Width           =   1485
         End
         Begin VB.Label Label73 
            Caption         =   "饮酒习惯"
            Height          =   255
            Left            =   -74760
            TabIndex        =   125
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label Label72 
            Caption         =   "病历（目前健康情况）"
            Height          =   315
            Left            =   -74760
            TabIndex        =   123
            Top             =   720
            Width           =   1845
         End
         Begin VB.Label Label71 
            Caption         =   "短期事业目标为何"
            Height          =   225
            Left            =   -74580
            TabIndex        =   119
            Top             =   5700
            Width           =   1575
         End
         Begin VB.Label Label70 
            Caption         =   "本客户长期事业目标为何"
            Height          =   285
            Left            =   -74580
            TabIndex        =   118
            Top             =   5190
            Width           =   2085
         End
         Begin VB.Label Label69 
            Caption         =   "客户对自己公司的态度"
            Height          =   255
            Left            =   -74580
            TabIndex        =   117
            Top             =   4740
            Width           =   2085
         End
         Begin VB.Label Label68 
            Caption         =   "本公司其他人员对客户的了解（何种关系，关系性质）"
            Height          =   555
            Left            =   -74580
            TabIndex        =   108
            Top             =   4140
            Width           =   2355
         End
         Begin VB.Label Label67 
            Caption         =   "原因"
            Height          =   285
            Left            =   -74580
            TabIndex        =   107
            Top             =   3750
            Width           =   1365
         End
         Begin VB.Label Label66 
            Caption         =   "本客户与本公司其他人关系是否良好"
            Height          =   345
            Left            =   -74580
            TabIndex        =   106
            Top             =   3300
            Width           =   1815
         End
         Begin VB.Label Label65 
            Caption         =   "在办公室有何“地位”象征"
            Height          =   315
            Left            =   -74580
            TabIndex        =   105
            Top             =   2880
            Width           =   2295
         End
         Begin VB.Line Line2 
            X1              =   -74970
            X2              =   -65190
            Y1              =   2790
            Y2              =   2790
         End
         Begin VB.Label Label64 
            Caption         =   "日  期"
            Height          =   195
            Left            =   -70920
            TabIndex        =   104
            Top             =   2430
            Width           =   1425
         End
         Begin VB.Label Label63 
            Caption         =   "职  衔"
            Height          =   315
            Left            =   -74580
            TabIndex        =   103
            Top             =   2430
            Width           =   1425
         End
         Begin VB.Label Label62 
            Caption         =   "在目前公司的前一个职衔"
            Height          =   195
            Left            =   -74580
            TabIndex        =   102
            Top             =   2160
            Width           =   2205
         End
         Begin VB.Line Line1 
            X1              =   -74970
            X2              =   -65220
            Y1              =   2070
            Y2              =   2070
         End
         Begin VB.Label Label61 
            Caption         =   "受雇职衔"
            Height          =   285
            Left            =   -69870
            TabIndex        =   101
            Top             =   1710
            Width           =   1305
         End
         Begin VB.Label Label60 
            Caption         =   "受雇时间"
            Height          =   315
            Left            =   -74580
            TabIndex        =   100
            Top             =   1710
            Width           =   1455
         End
         Begin VB.Label Label59 
            Caption         =   "公司地址"
            Height          =   285
            Left            =   -74580
            TabIndex        =   99
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label58 
            Caption         =   "公司名称"
            Height          =   285
            Left            =   -74580
            TabIndex        =   98
            Top             =   930
            Width           =   1425
         End
         Begin VB.Label Label57 
            Caption         =   "客户的前一个工作"
            Height          =   315
            Left            =   -74640
            TabIndex        =   97
            Top             =   660
            Width           =   1575
         End
         Begin VB.Label Label56 
            Caption         =   "子女喜好"
            Height          =   285
            Left            =   -74670
            TabIndex        =   88
            Top             =   5940
            Width           =   945
         End
         Begin VB.Label Label55 
            Caption         =   "子女教育"
            Height          =   285
            Left            =   -74670
            TabIndex        =   87
            Top             =   5370
            Width           =   1155
         End
         Begin VB.Label Label54 
            Caption         =   "是否有抚养权"
            Height          =   345
            Left            =   -70080
            TabIndex        =   86
            Top             =   4590
            Width           =   1515
         End
         Begin VB.Label Label53 
            Caption         =   "子女姓名、年龄"
            Height          =   285
            Left            =   -74670
            TabIndex        =   84
            Top             =   4650
            Width           =   1575
         End
         Begin VB.Label Label52 
            Caption         =   "结婚纪念日"
            Height          =   285
            Left            =   -74670
            TabIndex        =   83
            Top             =   3930
            Width           =   1965
         End
         Begin VB.Label Label51 
            Caption         =   "配偶兴趣/活动/社团"
            Height          =   405
            Left            =   -74670
            TabIndex        =   82
            Top             =   3090
            Width           =   1755
         End
         Begin VB.Label Label50 
            Caption         =   "配偶教育程度"
            Height          =   405
            Left            =   -74670
            TabIndex        =   81
            Top             =   2340
            Width           =   1395
         End
         Begin VB.Label Label49 
            Caption         =   "配偶姓名"
            Height          =   285
            Left            =   -74670
            TabIndex        =   80
            Top             =   1650
            Width           =   1215
         End
         Begin VB.Label Label48 
            Caption         =   "婚姻状况"
            Height          =   315
            Left            =   -74670
            TabIndex        =   79
            Top             =   930
            Width           =   1665
         End
         Begin VB.Label Label47 
            Caption         =   "课外活动、社团"
            Height          =   315
            Left            =   -74370
            TabIndex        =   76
            Top             =   5010
            Width           =   1455
         End
         Begin VB.Label Label46 
            Caption         =   "擅长运动是"
            Height          =   255
            Left            =   -74010
            TabIndex        =   74
            Top             =   3930
            Width           =   975
         End
         Begin VB.Label Label45 
            Caption         =   "大学所属学生会"
            Height          =   255
            Left            =   -74370
            TabIndex        =   72
            Top             =   3330
            Width           =   1335
         End
         Begin VB.Label Label44 
            Caption         =   "大学时期得奖记录"
            Height          =   255
            Left            =   -74550
            TabIndex        =   70
            Top             =   2640
            Width           =   1515
         End
         Begin VB.Label Label43 
            Caption         =   "学  位"
            Height          =   255
            Left            =   -69450
            TabIndex        =   67
            Top             =   2010
            Width           =   1065
         End
         Begin VB.Label Label42 
            Caption         =   "毕业日期"
            Height          =   315
            Left            =   -73860
            TabIndex        =   66
            Top             =   2010
            Width           =   1005
         End
         Begin VB.Label Label41 
            Caption         =   "大专名称"
            Height          =   225
            Left            =   -73860
            TabIndex        =   63
            Top             =   1530
            Width           =   945
         End
         Begin VB.Label Label40 
            Caption         =   "高中名称与就读期间"
            Height          =   345
            Left            =   -74730
            TabIndex        =   62
            Top             =   990
            Width           =   1665
         End
         Begin VB.Label Label25 
            Caption         =   "身体五官特征"
            Height          =   1335
            Left            =   480
            TabIndex        =   58
            Top             =   5190
            Width           =   285
         End
         Begin VB.Label Label23 
            Caption         =   "体  重"
            Height          =   285
            Left            =   3150
            TabIndex        =   57
            Top             =   4590
            Width           =   735
         End
         Begin VB.Label Label20 
            Caption         =   "身  高"
            Height          =   225
            Left            =   210
            TabIndex        =   56
            Top             =   4560
            Width           =   765
         End
         Begin VB.Label lblLhk 
            Caption         =   "籍    贯"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   6270
            TabIndex        =   55
            Top             =   4500
            Width           =   915
         End
         Begin VB.Label Label8 
            Caption         =   "性  别"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   3060
            TabIndex        =   53
            Top             =   3900
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "*"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   6
            Left            =   3750
            TabIndex        =   52
            Top             =   3870
            Width           =   135
         End
         Begin VB.Label Label10 
            Caption         =   "生  日"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   49
            Top             =   3900
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "邮件地址"
            Height          =   255
            Left            =   6270
            TabIndex        =   46
            Top             =   3915
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "*"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   4
            Left            =   7110
            TabIndex        =   45
            Top             =   3270
            Width           =   135
         End
         Begin VB.Label Label37 
            Caption         =   "手    机"
            ForeColor       =   &H00C00000&
            Height          =   225
            Index           =   0
            Left            =   6240
            TabIndex        =   44
            Top             =   3330
            Width           =   885
         End
         Begin VB.Label Label35 
            Caption         =   "单位地址"
            ForeColor       =   &H00C00000&
            Height          =   225
            Index           =   0
            Left            =   150
            TabIndex        =   42
            Top             =   1485
            Width           =   735
         End
         Begin VB.Label Label26 
            Caption         =   "电话(宅)"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   3360
            Width           =   825
         End
         Begin VB.Label Label24 
            Caption         =   "*"
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   900
            TabIndex        =   38
            Top             =   2670
            Width           =   165
         End
         Begin VB.Label lblgdate 
            Caption         =   "电话(公)"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   15
            Left            =   150
            TabIndex        =   37
            Top             =   2700
            Width           =   915
         End
         Begin VB.Label Label33 
            Caption         =   "家庭住址"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   35
            Top             =   2070
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "*"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   3
            Left            =   900
            TabIndex        =   33
            Top             =   1470
            Width           =   165
         End
         Begin VB.Label Label10 
            Caption         =   "职  务"
            ForeColor       =   &H00C00000&
            Height          =   225
            Index           =   2
            Left            =   6150
            TabIndex        =   32
            Top             =   900
            Width           =   555
         End
         Begin VB.Label Label1 
            Caption         =   "妮  称"
            Height          =   255
            Left            =   3060
            TabIndex        =   29
            Top             =   840
            Width           =   675
         End
         Begin VB.Label Label14 
            Caption         =   "*"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   5
            Left            =   900
            TabIndex        =   28
            Top             =   810
            Width           =   135
         End
         Begin VB.Label Label5 
            Caption         =   "姓  名"
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   27
            Top             =   810
            Width           =   615
         End
      End
      Begin VB.Frame frmGL 
         Caption         =   "管理楼盘信息:"
         Height          =   5895
         Left            =   -65760
         TabIndex        =   198
         Top             =   540
         Width           =   15225
         Begin VB.CommandButton cmdLQ 
            Caption         =   "清空"
            Height          =   315
            Left            =   11700
            TabIndex        =   290
            Top             =   4920
            Width           =   795
         End
         Begin VB.CommandButton cmdGx 
            Caption         =   "更新"
            Height          =   345
            Left            =   13440
            TabIndex        =   285
            Top             =   4890
            Width           =   765
         End
         Begin MSAdodcLib.Adodc adoLouPan 
            Height          =   375
            Left            =   9540
            Top             =   4170
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
         Begin VB.TextBox txtL 
            Height          =   465
            Index           =   0
            Left            =   1290
            TabIndex        =   205
            Tag             =   "30"
            Text            =   "0"
            Top             =   510
            Width           =   7395
         End
         Begin VB.TextBox txtL 
            Height          =   465
            Index           =   1
            Left            =   1290
            TabIndex        =   204
            Tag             =   "50"
            Text            =   "1"
            Top             =   1110
            Width           =   7395
         End
         Begin VB.TextBox txtL 
            Height          =   1125
            Index           =   2
            Left            =   1290
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   203
            Tag             =   "200"
            Text            =   "wbDN.frx":1478
            Top             =   1740
            Width           =   7395
         End
         Begin VB.TextBox txtL 
            Height          =   465
            Index           =   3
            Left            =   1290
            TabIndex        =   202
            Tag             =   "30"
            Text            =   "3"
            Top             =   3120
            Width           =   7395
         End
         Begin VB.TextBox txtL 
            Height          =   1605
            Index           =   4
            Left            =   1290
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   201
            Tag             =   "200"
            Text            =   "wbDN.frx":147C
            Top             =   3780
            Width           =   7395
         End
         Begin VB.CommandButton cmdLadd 
            Caption         =   "添加"
            Height          =   345
            Left            =   12540
            TabIndex        =   200
            Top             =   4890
            Width           =   855
         End
         Begin VB.CommandButton cmdLdel 
            Caption         =   "删除"
            Height          =   345
            Left            =   14220
            TabIndex        =   199
            Top             =   4890
            Width           =   825
         End
         Begin MSDataGridLib.DataGrid dtgLouPan 
            Bindings        =   "wbDN.frx":1480
            Height          =   4755
            Left            =   9000
            TabIndex        =   206
            Top             =   90
            Width           =   6225
            _ExtentX        =   10980
            _ExtentY        =   8387
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
               DataField       =   "w0"
               Caption         =   "楼盘名称"
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
               DataField       =   "w3"
               Caption         =   "保养单位"
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
                  ColumnWidth     =   2505.26
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   2505.26
               EndProperty
            EndProperty
         End
         Begin VB.Label Label32 
            Caption         =   "楼盘名称"
            Height          =   315
            Left            =   270
            TabIndex        =   211
            Top             =   570
            Width           =   1335
         End
         Begin VB.Label Label31 
            Caption         =   "地  址"
            Height          =   255
            Left            =   270
            TabIndex        =   210
            Top             =   1230
            Width           =   1515
         End
         Begin VB.Label Label30 
            Caption         =   "机组情况"
            Height          =   315
            Left            =   270
            TabIndex        =   209
            Top             =   1845
            Width           =   1545
         End
         Begin VB.Label Label29 
            Caption         =   "保养单位"
            Height          =   315
            Left            =   270
            TabIndex        =   208
            Top             =   3165
            Width           =   1425
         End
         Begin VB.Label Label28 
            Caption         =   "其它说明"
            Height          =   285
            Left            =   270
            TabIndex        =   207
            Top             =   3840
            Width           =   1005
         End
      End
      Begin VB.Frame frmJz 
         Height          =   5925
         Left            =   -74940
         TabIndex        =   197
         Top             =   1950
         Width           =   15225
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
            Height          =   3015
            Left            =   0
            TabIndex        =   297
            Top             =   2880
            Width           =   9165
            _ExtentX        =   16166
            _ExtentY        =   5318
            _Version        =   393216
            BackColor       =   15728356
            ForeColor       =   8404992
            Rows            =   15
            Cols            =   5
            FixedCols       =   0
            BackColorFixed  =   16777152
            ForeColorFixed  =   0
            BackColorBkg    =   15728356
            GridColorFixed  =   8404992
            GridColorUnpopulated=   8404992
            ScrollTrack     =   -1  'True
            SelectionMode   =   1
            AllowUserResizing=   1
            RowSizingMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgA 
            Height          =   2775
            Left            =   30
            TabIndex        =   277
            Top             =   120
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   4895
            _Version        =   393216
            BackColorBkg    =   -2147483634
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Frame frmJE 
            Caption         =   "机组信息编辑"
            Height          =   1875
            Left            =   0
            TabIndex        =   259
            Top             =   4080
            Width           =   8535
            Begin VB.CommandButton cmdEd 
               Caption         =   "编辑"
               Height          =   375
               Left            =   7740
               TabIndex        =   291
               Top             =   270
               Width           =   585
            End
            Begin VB.CommandButton cmdJgx 
               Caption         =   "更新"
               Height          =   345
               Left            =   7740
               TabIndex        =   276
               Top             =   1410
               Width           =   585
            End
            Begin VB.CommandButton cmdJadd 
               Caption         =   "添加"
               Height          =   375
               Left            =   7740
               TabIndex        =   275
               Top             =   630
               Width           =   585
            End
            Begin VB.CommandButton cmdJdel 
               Caption         =   "删除"
               Height          =   375
               Left            =   7740
               TabIndex        =   274
               Top             =   1020
               Width           =   585
            End
            Begin VB.TextBox txtYcou 
               Height          =   315
               Left            =   5250
               TabIndex        =   273
               Top             =   1380
               Width           =   2115
            End
            Begin VB.TextBox txtYj 
               Height          =   315
               Left            =   5250
               TabIndex        =   272
               Top             =   1027
               Width           =   2115
            End
            Begin VB.TextBox txtPjPz 
               Height          =   315
               Left            =   5250
               TabIndex        =   271
               Top             =   670
               Width           =   2115
            End
            Begin VB.TextBox txtJzCou 
               Height          =   315
               Left            =   1350
               TabIndex        =   267
               Top             =   1380
               Width           =   2535
            End
            Begin VB.TextBox txtJzxh 
               Height          =   315
               Left            =   1350
               TabIndex        =   265
               Top             =   1025
               Width           =   2535
            End
            Begin VB.TextBox txtJzPb 
               Height          =   315
               Left            =   1350
               TabIndex        =   264
               Top             =   670
               Width           =   2535
            End
            Begin VB.ComboBox comPz 
               Height          =   300
               ItemData        =   "wbDN.frx":1498
               Left            =   1350
               List            =   "wbDN.frx":14B1
               TabIndex        =   260
               Top             =   330
               Width           =   2535
            End
            Begin VB.Label lblJid 
               Caption         =   "lblJid"
               Height          =   255
               Left            =   5340
               TabIndex        =   278
               Top             =   300
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.Label Label88 
               Caption         =   "配件数量"
               Height          =   285
               Left            =   4350
               TabIndex        =   270
               Top             =   1440
               Width           =   765
            End
            Begin VB.Label Label87 
               Caption         =   "品牌型号"
               Height          =   315
               Left            =   4320
               TabIndex        =   269
               Top             =   1080
               Width           =   825
            End
            Begin VB.Label Label86 
               Caption         =   "配件种类"
               Height          =   255
               Left            =   4350
               TabIndex        =   268
               Top             =   720
               Width           =   765
            End
            Begin VB.Label Label85 
               Caption         =   "机组数量"
               Height          =   225
               Left            =   390
               TabIndex        =   266
               Top             =   1440
               Width           =   825
            End
            Begin VB.Label Label39 
               Caption         =   "机组型号"
               Height          =   255
               Left            =   390
               TabIndex        =   263
               Top             =   1110
               Width           =   855
            End
            Begin VB.Label Label38 
               Caption         =   "机组品牌"
               Height          =   255
               Left            =   390
               TabIndex        =   262
               Top             =   750
               Width           =   885
            End
            Begin VB.Label Label36 
               Caption         =   "设备类型:"
               Height          =   285
               Left            =   390
               TabIndex        =   261
               Top             =   390
               Width           =   1305
            End
         End
         Begin VB.Label lblTX 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   9240
            TabIndex        =   299
            Top             =   4590
            Width           =   5475
         End
      End
      Begin VB.Label lblKid 
         Caption         =   "lblKid"
         Height          =   255
         Left            =   -74850
         TabIndex        =   215
         Top             =   210
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   -73950
         TabIndex        =   18
         Top             =   960
         Width           =   165
      End
      Begin VB.Label Label17 
         Caption         =   "注：*为必填项目"
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   7890
         TabIndex        =   17
         Top             =   -960
         Width           =   1875
      End
      Begin VB.Label Label14 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   0
         Left            =   -73950
         TabIndex        =   16
         Top             =   480
         Width           =   165
      End
      Begin VB.Label Label6 
         Caption         =   "国税号"
         Height          =   285
         Left            =   -64320
         TabIndex        =   15
         Top             =   510
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "开户银行"
         Height          =   285
         Left            =   -69300
         TabIndex        =   14
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label8 
         Caption         =   "账  号"
         Height          =   285
         Index           =   2
         Left            =   -64320
         TabIndex        =   13
         Top             =   982
         Width           =   555
      End
      Begin VB.Label Label12 
         Caption         =   "所属区域"
         Height          =   285
         Left            =   -69300
         TabIndex        =   7
         Top             =   510
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "地  址"
         Height          =   315
         Left            =   -64320
         TabIndex        =   6
         Top             =   1455
         Width           =   585
      End
      Begin VB.Label Label13 
         Caption         =   "行业性质"
         Height          =   285
         Left            =   -69300
         TabIndex        =   4
         Top             =   1455
         Width           =   1005
      End
      Begin VB.Label lblgdate 
         Caption         =   "客户代码"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   3
         Top             =   990
         Width           =   975
      End
      Begin VB.Label lblgdate 
         Caption         =   $"wbDN.frx":14FF
         Height          =   375
         Index           =   4
         Left            =   -74760
         TabIndex        =   2
         Top             =   1455
         Width           =   975
      End
      Begin VB.Label lblgdate 
         Caption         =   "客户全称"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   1
         Top             =   510
         Width           =   975
      End
   End
   Begin VB.TextBox txtQrq 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4470
      TabIndex        =   241
      Top             =   570
      Width           =   1455
   End
   Begin VB.Label lblKuid 
      Caption         =   "lblKuid"
      Height          =   255
      Left            =   10830
      TabIndex        =   288
      Top             =   990
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblKywy 
      Caption         =   "lblKywy"
      Height          =   285
      Left            =   9330
      TabIndex        =   287
      Top             =   960
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblXmPd 
      Caption         =   "Label36"
      Height          =   225
      Left            =   7950
      TabIndex        =   258
      Top             =   1020
      Width           =   1065
   End
   Begin VB.Label Label22 
      Caption         =   "项目平台"
      Height          =   195
      Left            =   7080
      TabIndex        =   257
      Top             =   1050
      Width           =   765
   End
   Begin VB.Label Label3 
      Caption         =   "地  址"
      Height          =   315
      Left            =   390
      TabIndex        =   254
      Top             =   990
      Width           =   585
   End
   Begin VB.Label lblWy 
      Caption         =   "lblWy"
      Height          =   195
      Left            =   10080
      TabIndex        =   252
      Top             =   600
      Width           =   2265
   End
   Begin VB.Label lblYz 
      Caption         =   "lblYz"
      Height          =   195
      Left            =   10080
      TabIndex        =   251
      Top             =   210
      Width           =   2235
   End
   Begin VB.Label lblygFy 
      Caption         =   "ygFy"
      Height          =   285
      Left            =   7950
      TabIndex        =   250
      Top             =   630
      Width           =   1065
   End
   Begin VB.Label Label7 
      Caption         =   "已归合同项目费用"
      Height          =   315
      Left            =   6360
      TabIndex        =   249
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label lblXmfy 
      Caption         =   "xmFy"
      Height          =   345
      Left            =   7950
      TabIndex        =   248
      Top             =   150
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "项目总费用"
      Height          =   315
      Left            =   6900
      TabIndex        =   247
      Top             =   150
      Width           =   945
   End
   Begin VB.Label lblXid 
      Caption         =   "lblXid"
      Height          =   255
      Left            =   1290
      TabIndex        =   243
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblgdate 
      Caption         =   "初次合同签订期"
      Height          =   255
      Index           =   23
      Left            =   2940
      TabIndex        =   242
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label34 
      Caption         =   "项目代码"
      Height          =   195
      Left            =   210
      TabIndex        =   240
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label14 
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   1020
      TabIndex        =   239
      Top             =   120
      Width           =   165
   End
   Begin VB.Label Label27 
      Caption         =   "项目名称"
      Height          =   285
      Left            =   210
      TabIndex        =   238
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "wbDN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public adoHy As Object
Public adoQy As Object
Public modFi As Boolean '是否修改过
Public adoLxr As Object '联系人详细信息
Public khAdd As Boolean '是否为新建客户
Public adoA As Object '机组的ADO
Public adoKhmc As Object '客户名称框的ado
Dim timZm As Integer '1审核 2新审核
Dim LName As String
Public Lc As Integer
Public LCRen As String
Public LCUid As String
Public Fwid As String


Public Sub dtgPFF()
Dim oo As Integer
For oo = 1 To dtgP.Rows - 1
    dtgP.RowHeight(oo) = dtgP.RowHeight(0) * 2
Next
dtgP.Clear
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "日期": dtgP.Col = 1: dtgP.Text = "姓名": dtgP.Col = 2: dtgP.Text = "职能": dtgP.Col = 3: dtgP.Text = "评审建议": dtgP.Col = 4: dtgP.Text = "审核":
dtgP.ColWidth(0) = 1665
dtgP.ColWidth(1) = 1005
dtgP.ColWidth(2) = 0
 dtgP.ColWidth(3) = 4290: dtgP.ColWidth(4) = 1035
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub




Private Sub cmdBack_Click()
On Error Resume Next
Dim tt As String
'khAdd.Close
Dim ii As Integer
'If modFi = True Then
'    ii = MsgBox("退出将不保存数据！", vbYesNo + vbInformation + vbDefaultButton2, "请确认")
'    If ii = vbNo Then
'    Exit Sub
'    End If
'End If
If txtXMMC.Text = "" Then '判断项目名称
    ii = MsgBox("项目名称不能为空,退出将删除此项目!", vbInformation + vbYesNo, "询问")
    If ii = vbYes Then
        tt = "delete from xmzl where xid=" & Val(lblXid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Else
        Exit Sub
    End If
End If

If txtKhmc.Text = "" Then
'    ii = MsgBox("没有输入项目名称或客户名称,退出将不保存数据！", vbYesNo + vbInformation + vbDefaultButton2, "请确认")
'    If ii = vbNo Then
'        Exit Sub
'    Else
        tt = "delete from khzl where kid=" & Val(lblKid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

        '删除流程签字表中的记录
        tt = "delete from qmrz where qdbh='" & lblKid.Caption & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'End If
End If




Call mod1.DelDKZ ' '退出表单时删除打开记录,以让别人能打开此单据
wbDN.Visible = False
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf frmKhBr.Visible = True Then
    
    frmKhBr.Show
    If wbDN.khAdd = True Then
        frmKhBr.tabCx.Tab = 0
        tt = "vXmNew('" & mod1.DName & "','" & mod1.DHid & "')"
        frmKhBr.adoKhBr.Close
        frmKhBr.adoKhBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        Set frmKhBr.dtgKh.DataSource = frmKhBr.adoKhBr
    End If
    frmKhBr.Enabled = True
    frmKhBr.ZOrder 0
ElseIf frmKhbrG.Visible = True Then
    
    frmKhbrG.Enabled = True
    frmKhbrG.ZOrder 0
ElseIf frmGzNr.Visible = True Then
    frmGzNr.Enabled = True
    frmGzNr.ZOrder 0

ElseIf frmWbNew.Visible = True Then
    frmWbNew.Enabled = True
    frmWbNew.ZOrder 0
ElseIf FMXC.Visible = True Then
    FMXC.Enabled = True
    FMXC.ZOrder 0
ElseIf FmxcNew.Visible = True Then
    FmxcNew.Enabled = True
    FmxcNew.ZOrder 0
End If
frmHyxz.Visible = False
End Sub

Private Sub cmdBackA_Click()
On Error Resume Next
Dim tt As String
'khAdd.Close
Dim ii As Integer
If cmdSave.Enabled = True Or cmdSave.Enabled = True Then
    ii = MsgBox("退出将不保存数据！", vbYesNo + vbInformation + vbDefaultButton2, "请确认")
    If ii = vbNo Then
    Exit Sub
    End If
End If

wbDN.Visible = False
End Sub

Private Sub cmdBackB_Click()
On Error Resume Next
Dim tt As String
'khAdd.Close
Dim ii As Integer
If cmdSave.Enabled = True Or cmdSave.Enabled = True Then
    ii = MsgBox("退出将不保存数据！", vbYesNo + vbInformation + vbDefaultButton2, "请确认")
    If ii = vbNo Then
    Exit Sub
    End If
End If

wbDN.Visible = False
End Sub

Private Sub cmdDel_Click()
Dim ii As Integer
ii = MsgBox("您确认要取消所有修改吗?", vbExclamation + vbYesNo + vbDefaultButton2, "请确认")
If ii = vbYes Then



End If
End Sub





Private Sub cmdDing1_Click()
If txtQM.Text = "" And optT2.Value = True Then
    MsgBox "请告诉拒绝的理由先！"
    Exit Sub
End If
timZm = 1 '审核
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "项目资料"
    mod1.cmd.Parameters("@NBLX") = "审核"
    mod1.cmd.Parameters("@bh") = Trim(txtKhDm.Text)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtMan.Text)
    mod1.cmd.Parameters("@mt2") = Trim(lblRid.Caption)
    mod1.cmd.Parameters("@mt3") = mod1.Bm
    mod1.cmd.Parameters("@mt4") = mod1.Qy
    mod1.cmd.Parameters("@mt5") = mod1.GJId
    mod1.cmd.Parameters("@mt6") = mod1.GJR
    mod1.cmd.Parameters("@mt7") = lblQM(Val(lblLc.Caption) - 1).Caption

    mod1.cmd.Parameters("@mlt1") = txtQM.Text

    mod1.cmd.Parameters("@mm1") = Val(lblLc.Caption)
    mod1.cmd.Parameters("@mm2") = Val(lblFwid.Caption)
    mod1.cmd.Parameters("@mm3") = mod1.BTZ '业务属性

        If OptT1.Value = True Then
            mod1.cmd.Parameters("@mb1") = 1
        Else
            mod1.cmd.Parameters("@mb1") = 0
        End If

    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If
End Sub

Private Sub cmdDing_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next

If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

timZm = 2 '签字
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "项目资料"
    mod1.cmd.Parameters("@NBLX") = "新审核"
    mod1.cmd.Parameters("@bh") = lblXid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtXMMC.Text
    mod1.cmd.Parameters("@mt11") = lblYwy.Caption
    mod1.cmd.Parameters("@mt12") = lblUid.Caption
    mod1.cmd.Parameters("@mlt1") = txtQM.Text
    mod1.cmd.Parameters("@mm1") = Lc
    mod1.cmd.Parameters("@mm2") = Fwid
    If OptT1.Value = True Then
        mod1.cmd.Parameters("@mb1") = 1 '同意
    Else
        mod1.cmd.Parameters("@mb1") = 0 '拒绝
    End If
    mod1.cmd.Parameters("@md1") = Null
     Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        cmdDing.Enabled = False
    
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
frmQm.Visible = False
End Sub


Private Sub cmdED_Click()
On Error Resume Next
If adoA.RecordCount = 0 Then Exit Sub
lblJid.Caption = ""
dtgA.Col = 10
lblJid.Caption = dtgA.Text
If Val(lblJid.Caption) = 0 Then
    MsgBox "请选择上方一条机组信息记录,再进行编辑!"
    Exit Sub
End If

comPz.Text = ""
txtJzPb.Text = ""
txtJzCou.Text = ""
txtJzxh.Text = ""
txtPjPz.Text = ""
txtYJ.Text = ""
txtYcou.Text = ""

dtgA.Col = 1
comPz.Text = dtgA.Text
dtgA.Col = 2
txtJzPb.Text = dtgA.Text
dtgA.Col = 3
txtJzxh.Text = dtgA.Text
dtgA.Col = 4
txtJzCou.Text = dtgA.Text
dtgA.Col = 5
txtPjPz.Text = dtgA.Text
dtgA.Col = 6
txtYJ.Text = dtgA.Text
dtgA.Col = 7
txtYcou.Text = dtgA.Text

End Sub

Private Sub cmdGx_Click()

On Error Resume Next

Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "lopanGx"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@kid") = Val(lblKid.Caption)
    mod1.cmd.Parameters("@w0") = Trim(txtL(0).Text)
    mod1.cmd.Parameters("@w1") = Trim(txtL(1).Text)
    mod1.cmd.Parameters("@w2") = Trim(txtL(2).Text)
    mod1.cmd.Parameters("@w3") = Trim(txtL(3).Text)
    mod1.cmd.Parameters("@w4") = Trim(txtL(4).Text)
    mod1.cmd.Parameters("@lid") = wbDN.adoLouPan.Recordset.Fields("lid").Value
    mod1.cmd.Execute
    
    Set cmd = Nothing
    wbDN.adoLouPan.Recordset.Requery
    Set dtgLouPan.DataSource = adoLouPan
End Sub

Private Sub cmdJadd_Click()
On Error Resume Next
Dim tt As String
Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "jzAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@pz") = Trim(comPz.Text)
    mod1.cmd.Parameters("@jzpb") = Trim(txtJzPb.Text)
    mod1.cmd.Parameters("@jzcou") = Val(txtJzCou.Text)
    mod1.cmd.Parameters("@jzxh") = Trim(txtJzxh.Text)
    mod1.cmd.Parameters("@pjpz") = Trim(txtPjPz.Text)
    mod1.cmd.Parameters("@yj") = Trim(txtYJ.Text)
    mod1.cmd.Parameters("@ycou") = Val(txtYcou.Text)
    mod1.cmd.Parameters("@khdh") = Trim(txtKhDm.Text)
    mod1.cmd.Parameters("@kid") = Trim(lblKid.Caption)
    mod1.cmd.Parameters("@mch") = ""
    mod1.cmd.Execute
    tt = mod1.cmd.Parameters("@mch").Value
    Set cmd = Nothing
    If tt = "" Then
        MsgBox "网络出现故障,请再试一次,如果还是提交不成功,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else
        Call mod1.khJBound(lblKid.Caption)
        
        comPz.Text = ""
        txtJzPb.Text = ""
        txtJzCou.Text = ""
        txtJzxh.Text = ""
        txtPjPz.Text = ""
        txtYJ.Text = ""
        txtYcou.Text = ""
        lblJid.Caption = ""
    End If
End Sub

Private Sub cmdJdel_Click()
Dim ii As Integer
On Error Resume Next
ii = MsgBox("是否删除此条记录?", vbInformation + vbYesNo, "询问")
If ii = vbYes Then
    Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "jzDel"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@jid") = Val(lblJid.Caption)
        mod1.cmd.Execute
    Set cmd = Nothing
    
    Call mod1.khJBound(lblKid.Caption)
    
    comPz.Text = ""
    txtJzPb.Text = ""
    txtJzCou.Text = ""
    txtJzxh.Text = ""
    txtPjPz.Text = ""
    txtYJ.Text = ""
    txtYcou.Text = ""
    lblJid.Caption = ""
End If
End Sub

Private Sub cmdJgx_Click()
On Error Resume Next

Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "jzUpdate"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@pz") = Trim(comPz.Text)
    mod1.cmd.Parameters("@jzpb") = Trim(txtJzPb.Text)
    mod1.cmd.Parameters("@jzcou") = Val(txtJzCou.Text)
    mod1.cmd.Parameters("@jzxh") = Trim(txtJzxh.Text)
    mod1.cmd.Parameters("@pjpz") = Trim(txtPjPz.Text)
    mod1.cmd.Parameters("@yj") = Trim(txtYJ.Text)
    mod1.cmd.Parameters("@ycou") = Val(txtYcou.Text)
    mod1.cmd.Parameters("@jid") = Val(lblJid.Caption)
    mod1.cmd.Execute
Set cmd = Nothing

Call mod1.khJBound(lblKid.Caption)

comPz.Text = ""
txtJzPb.Text = ""
txtJzCou.Text = ""
txtJzxh.Text = ""
txtPjPz.Text = ""
txtYJ.Text = ""
txtYcou.Text = ""
lblJid.Caption = ""
End Sub

Private Sub cmdLadd_Click()
On Error Resume Next

Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "lopanAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@kid") = Val(lblKid.Caption)
    mod1.cmd.Parameters("@w0") = Trim(txtL(0).Text)
    mod1.cmd.Parameters("@w1") = Trim(txtL(1).Text)
    mod1.cmd.Parameters("@w2") = Trim(txtL(2).Text)
    mod1.cmd.Parameters("@w3") = Trim(txtL(3).Text)
    mod1.cmd.Parameters("@w4") = Trim(txtL(4).Text)
    mod1.cmd.Execute
    
    Set cmd = Nothing
    wbDN.adoLouPan.Recordset.Requery
    Set dtgLouPan.DataSource = adoLouPan
    
'    wbDN.adoLouPan.Recordset.AddNew "kid", Val(lblKid.Caption)
'     For oo = 0 To 4
'        txtL(oo).Locked = False
'        wbDN.adoLouPan.Recordset.Update "w" & oo, txtL(oo).Text
'     Next







End Sub

Private Sub cmdLDel_Click()
Dim oo As Integer
Dim tt As String
On Error Resume Next
Dim ii As Integer
ii = MsgBox("是否删除此条记录?", vbInformation + vbYesNo, "询问")
If ii = vbYes Then
    tt = "delete from khlopan where lid=" & adoLouPan.Recordset.Fields("lid").Value
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    mod1.HTP.Close
    For oo = 0 To 4
        txtL(oo).Text = ""
    Next
    adoLouPan.Recordset.Requery
    Set dtgLouPan.DataSource = adoLouPan
End If
End Sub

Private Sub cmdLeft_Click()
Dim tt As String
Dim oo As Integer
On Error Resume Next
    frmWait.Show
    frmWait.ZOrder 0
    frmWait.Refresh
    frmWait.faWait.Play
    adoRen.Recordset.MovePrevious
    cmdRight.Enabled = True
'tt = "vkhren2(" & wbDN.adoRen.Recordset.Fields("rid").Value & ")"
'wbDN.adoLxr.Close
'wbDN.adoLxr.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
'
'wbDN.txtMan.Text = wbDN.adoLxr.Fields("khMan").Value '联系人
'wbDN.txtHk.Text = wbDN.adoLxr.Fields("khHk").Value '户口
'
' wbDN.lblXb.Caption = wbDN.adoLxr.Fields("khSex").Value '性别
' If wbDN.lblXb.Caption = "男" Then
'    wbDN.optMan.Value = True
' ElseIf wbDN.lblXb.Caption = "女" Then
'    wbDN.optWoman.Value = True
' End If
'wbDN.dtpSr.Value = wbDN.adoLxr.Fields("khSr").Value '生日
'wbDN.txtZw.Text = wbDN.adoLxr.Fields("khZw").Value '职务
'wbDN.txtLpho.Text = wbDN.adoLxr.Fields("khDpho").Value '电话
'wbDN.txtLdwdz.Text = wbDN.adoLxr.Fields("khDwadr").Value '单位地址
'wbDN.txtLjpho.Text = wbDN.adoLxr.Fields("khJpho").Value '家庭电话
'wbDN.txtLjmob.Text = wbDN.adoLxr.Fields("khMob").Value '手机
'wbDN.txtLjadr.Text = wbDN.adoLxr.Fields("khJadr").Value '家庭地址
'
'For oo = 0 To 82
'    wbDN.txtK(oo).Text = wbDN.adoLxr.Fields("kh" & oo).Value
'Next
'wbDN.lblYwy.Caption = wbDN.adoLxr.Fields("ywy").Value
'wbDN.lblUid.Caption = wbDN.adoLxr.Fields("uid").Value
'wbDN.lblXywy.Caption = wbDN.adoLxr.Fields("xywy").Value
'wbDN.lblXuid.Caption = wbDN.adoLxr.Fields("xuid").Value
'wbDN.lblLc.Caption = wbDN.adoLxr.Fields("lc").Value
'wbDN.lblLcRen.Caption = wbDN.adoLxr.Fields("lcRen").Value
'wbDN.lblLcUid.Caption = wbDN.adoLxr.Fields("lcUid").Value
'wbDN.lblFwid.Caption = wbDN.adoLxr.Fields("Fwid").Value
'
''更新签字
'Call mod1.OpenKHAN
'frmWait.Visible = False

'adoRen.Recordset.MovePrevious
'If adoRen.Recordset.BOF = True Then
'    cmdLeft.Enabled = False
'End If
'adoRen.Recordset.MoveNext
If adoRen.Recordset.BOF = True Then
    'cmdLeft.Enabled = False
    adoRen.Recordset.MoveLast
End If
End Sub

Private Sub cmdOK_Click()
'Dim tt As String
'On Error Resume Next
'
''tt = "Select * from khRen where khMan='" & txtMan.Text & "'"
''wbDN.adoRen.Recordset.Close
''wbDN.adoRen.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'
'If txtMan.Text <> "" Then
'wbDN.adoRen.Recordset.Update "khdh", txtDm.Text '客户代号
'wbDN.adoRen.Recordset.Update "khMan", txtMan.Text '联系人
'wbDN.adoRen.Recordset.Update "khHk", txtHk.Text '户口
' '性别
' If wbDN.optMan.Value = True Then
'    wbDN.lblXb.Caption = "男"
' ElseIf wbDN.optWoman.Value = True Then
'    wbDN.lblXb.Caption = "女"
' End If
'wbDN.adoRen.Recordset.Update "khSex", wbDN.lblXb.Caption
'wbDN.adoRen.Recordset.Update "khWhcd", wbDN.comWhcd.Text '文化程度
'wbDN.adoRen.Recordset.Update "khOld", wbDN.txtOld.Text '年龄
'wbDN.adoRen.Recordset.Update "khczmm", wbDN.comCzmm.Text '政治面貌
'wbDN.adoRen.Recordset.Update "khSr", wbDN.dtpSr.Value '生日
'wbDN.adoRen.Recordset.Update "khHf", wbDN.comHf.Text '婚否
'wbDN.adoRen.Recordset.Update "khDw", wbDN.txtDw.Text '单位
'wbDN.adoRen.Recordset.Update "khZw", wbDN.txtZw.Text '职务
'wbDN.adoRen.Recordset.Update "khYb", wbDN.txtLYb.Text '邮编
'wbDN.adoRen.Recordset.Update "khDpho", wbDN.txtLpho.Text '电话
'wbDN.adoRen.Recordset.Update "khCz", wbDN.txtLfax.Text '传真
'wbDN.adoRen.Recordset.Update "khDwadr", wbDN.txtLdwdz.Text '单位地址
'wbDN.adoRen.Recordset.Update "khJpho", wbDN.txtLjpho.Text '家庭电话
'wbDN.adoRen.Recordset.Update "khMob", wbDN.txtLjmob.Text '手机
'wbDN.adoRen.Recordset.Update "khJyb", wbDN.txtLjyb.Text '邮编
'wbDN.adoRen.Recordset.Update "khJadr", wbDN.txtLjadr.Text '家庭地址
'wbDN.adoRen.Recordset.Update "khJtzk", wbDN.txtLjzk.Text '家庭状况
'wbDN.adoRen.Recordset.Update "khXg", wbDN.txtXg.Text '性格
'wbDN.adoRen.Recordset.Update "khSh", wbDN.txtSh.Text '嗜好
'wbDN.adoRen.Recordset.Update "khYd", wbDN.txtLyd.Text '优点
'wbDN.adoRen.Recordset.Update "khQd", wbDN.txtLqd.Text '缺点
'wbDN.adoRen.Recordset.UpdateBatch
'
'wbDN.cmdRadd.Enabled = True
'End If
End Sub

Private Sub cmdLQx_Click()

End Sub

Private Sub cmdLQ_Click()
txtL(0).Text = ""
txtL(1).Text = ""
txtL(2).Text = ""
txtL(3).Text = ""
txtL(4).Text = ""
End Sub

Private Sub cmdMod_Click()

If Val(lblXmPd.Caption) >= 60 And mod1.KhK = 0 And mod1.Bm <> "维销部2" Then '如果是业务员,在项目到60平台后,不能随便修改数据.
    MsgBox "此项目已经接近成功,不能擅自修改数据,如要修改,请与您的销售经理联系!"
    Exit Sub
End If
Call mod1.XmKhUnLocked
cmdSave.Enabled = True
modFi = True
If Val(lblLc.Caption) = 2 Then
    cmdRadd.Enabled = True
    cmdRdel.Enabled = True
Else
    cmdRadd.Enabled = False
    cmdRdel.Enabled = False
End If
tabRen.Enabled = True
dtgP.Visible = False

End Sub

Private Sub cmdNew_Click()
Dim tt As String
Dim oo As Integer
On Error Resume Next
'frmRen.Enabled = True


tt = "select lcou from newRenLc where nlb=38"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText


    '添加新联系人
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "khRenAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@khdh") = wbDN.txtKhDm.Text
    mod1.cmd.Parameters("@xid") = Val(lblXid.Caption)
    mod1.cmd.Parameters("@Lcou") = mod1.HTP.Fields("lcou").Value  '流程总数
    mod1.cmd.Parameters("@Lcou") = 3
    mod1.cmd.Parameters("@Lc") = 0 '当前流程
    mod1.cmd.Parameters("@lcRen") = mod1.DName
    mod1.cmd.Parameters("@lcUid") = mod1.DHid
    mod1.cmd.Parameters("@nLb") = 38
    mod1.cmd.Parameters("@kid") = lblKid.Caption
    mod1.cmd.Parameters("@rid") = 0
    mod1.cmd.Execute
    lblRid.Caption = mod1.cmd.Parameters("@rid").Value
    Set mod1.cmd = Nothing
    '设置流程按钮
    Call mod1.khLcBut(38)

    cmdLeft.Enabled = False
    cmdRight.Enabled = False

    'dtgRen.Visible = False
    wbDN.tabRen.Enabled = True
    
        tt = "select * from khren where rid=" & Val(lblRid.Caption)
        adoLxr.Close
        adoLxr.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        wbDN.adoLxr.Update "khdh", txtKhDm.Text '客户代号
        wbDN.adoLxr.Update "khMan", txtMan.Text '联系人
        wbDN.adoLxr.Update "khHk", txtHk.Text '户口
         '性别
         If wbDN.optMan.Value = True Then
            wbDN.lblXb.Caption = "男"
         ElseIf wbDN.optWoman.Value = True Then
            wbDN.lblXb.Caption = "女"
         End If
        wbDN.adoLxr.Update "khSex", wbDN.lblXb.Caption
        wbDN.adoLxr.Update "khSr", DateSerial(Year(dtpSr.Value), Month(dtpSr.Value), Day(dtpSr.Value))  '生日
        wbDN.adoLxr.Update "khZw", wbDN.txtZw.Text '职务
        wbDN.adoLxr.Update "khDpho", wbDN.txtLpho.Text '电话
        wbDN.adoLxr.Update "khDwadr", wbDN.txtLdwdz.Text '单位地址
        wbDN.adoLxr.Update "khJpho", wbDN.txtLjpho.Text '家庭电话
        wbDN.adoLxr.Update "khMob", wbDN.txtLjmob.Text '手机
        wbDN.adoLxr.Update "khJadr", wbDN.txtLjadr.Text '家庭地址
        For oo = 0 To 82
            wbDN.adoLxr.Update "kh" & oo, wbDN.txtK(oo).Text
        Next
        wbDN.adoLxr.Update "lc", 1
        wbDN.adoLxr.Update "lcRen", mod1.DName
        wbDN.adoLxr.Update "lcUid", mod1.DHid
        wbDN.adoLxr.Update "kid", Val(wbDN.lblKid.Caption)
        wbDN.adoLxr.UpdateBatch
    cmdLeft.Enabled = True
    cmdRight.Enabled = True
    '更新联系人列表
    tt = "vkhren1(" & Val(wbDN.lblKid.Caption) & ",'" & wbDN.lblYwy.Caption & "'," & Val(wbDN.lblXid.Caption) & ")"
    wbDN.adoRen.Recordset.Close
    wbDN.adoRen.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    Set wbDN.dtgRen.DataSource = wbDN.adoRen
    cmdNew.Enabled = True
    dtgRen.Visible = True
    cmdRdel.Enabled = True
    
    'If lblFwid.Caption = "" Then
'''''''''        '添加事务
'''''''''        Call mod1.EnventAdd("项目资料", txtMan.Text, mod1.DName, mod1.DHid, lblRid.Caption, lblQM(0).Caption, "", "", mod1.DName, mod1.DHid, 0, lblRid.Caption)
    'End If
    lblLcRen.Caption = mod1.DName
    lblLcUid.Caption = mod1.DHid
    Call mod1.khRQing
    cmdQm(0).Enabled = True
    cmdQm(1).Enabled = True
End Sub

Private Sub cmdNQ_Click()
Dim ii As Integer
Dim oo As Integer
On Error Resume Next


If lblTX.Caption = "审核完毕!" Then Exit Sub
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

If LCUid <> mod1.DHid Then
        MsgBox "此处应由" & LCRen & "签字! 请您不要再点"
        Exit Sub
End If

frmQm.Visible = True
If Lc = 1 Then
    optT2.Enabled = False
    OptT1.Value = True
    
Else
    OptT1.Enabled = True
    optT2.Enabled = True
    OptT1.Value = False
    optT2.Value = False
End If
End Sub

Private Sub cmdQing_Click()
    Call mod1.khRQing
    cmdNew.Enabled = True
    '不能签字
    cmdQm(0).Enabled = False
    cmdQm(1).Enabled = False
    cmdQm(0).Caption = ""
    lblTm(0).Caption = ""
    cmdQm(1).Caption = ""
    lblTm(1).Caption = ""
End Sub


Private Sub cmdQm_Click(Index As Integer)

If txtMan.Text = "" Or txtZw.Text = "" Or txtLpho.Text = "" Or txtLjmob.Text = "" Or (optMan.Value = False And optWoman.Value = False) Then
    MsgBox "客户资料不完整，请正确填写带星号的内容！"
    Exit Sub
End If
If cmdQm(Index).Caption <> "" Then Exit Sub

If Index = 0 And cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If


    If Index + 1 <> lblLc.Caption Then '不能在不相干的位置上乱点
    
        Exit Sub
    End If
    If mod1.BmJl = False Then
        If lblLcUid.Caption <> mod1.DHid Then
            MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
            Exit Sub
        End If
    End If
frmQm.Visible = True
OptT1.Value = False
optT2.Value = False
OptT1.Enabled = True
optT2.Enabled = True
txtQM.Text = ""
Exit Sub
Dim tt As String
Dim Tywy As String '单子流转到下一人的姓名
Dim Tuid As String
Dim Oywy As String '原来流转人的名字
Dim Ouid As String '原来流转人的工号

On Error Resume Next

Oywy = lblLcRen.Caption
Ouid = lblLcUid.Caption






Dim Zi As Integer
Zi = MsgBox("是否确认签字?", vbYesNo)
If Zi = vbNo Then Exit Sub

    lblLc.Caption = lblLc.Caption + 1

    
    '更新表khRen中的lcRen,lcUid 字段,以及QMRZ表中的相应字段.
                Set mod1.cmd = CreateObject("adodb.command")
                mod1.cmd.ActiveConnection = mod1.cc
                mod1.cmd.CommandText = "QMRZQM"
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@nlb") = 38 '单子(报销单)种类
                mod1.cmd.Parameters("@lc") = lblLc.Caption  '当前流程
                mod1.cmd.Parameters("@Dname") = mod1.DName
                mod1.cmd.Parameters("@uid") = mod1.DHid
                mod1.cmd.Parameters("@btz") = mod1.BTZ '业务属性
                mod1.cmd.Parameters("@zid") = cmdQm(Index).Tag '流程顺序
                mod1.cmd.Parameters("@Qdbh") = lblRid.Caption  '联系人编号
                mod1.cmd.Parameters("@pje") = ""   '评审建议
                mod1.cmd.Parameters("@bm") = mod1.Bm
                mod1.cmd.Parameters("@qy") = mod1.Qy
                mod1.cmd.Parameters("@Gren") = mod1.GJR
                mod1.cmd.Parameters("@Guid") = mod1.GJId
                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
                mod1.cmd.Parameters("@yid") = lblUid.Caption
                mod1.cmd.Parameters("@comid") = mod1.comId
                mod1.cmd.Execute
                Tywy = mod1.cmd.Parameters("@Tywy").Value
                Tuid = mod1.cmd.Parameters("@Tuid").Value
                Set cmd = Nothing
                cmdQm(Index).Caption = mod1.DName
                lblTm(Index).Caption = mod1.DQda
                lblLcRen.Caption = Tywy
                lblLcUid.Caption = Tuid
                

If lblQM(Index).Caption = "销售经理" Then
    Call mod1.EnventFinish(wbDN.lblFwid.Caption)
    tt = "update khren set Pwf=1 where xid=" & Val(lblXid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    MsgBox "您完成了对该业务员的审核."
Else
    '添加事务
    Call mod1.EnventAdd("项目资料", txtMan.Text, lblLcRen.Caption, lblLcUid.Caption, lblRid.Caption, lblQM(Index + 1).Caption, Oywy, Ouid, lblYwy.Caption, lblUid.Caption, lblFwid.Caption, lblRid.Caption)
    MsgBox "现在,此联系人资料将交由 " & Tywy & " 来审阅!"
End If

If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent(1)
End If
End Sub

Private Sub cmdQX_Click()
Dim ii As Integer
On Error Resume Next
ii = MsgBox("确定取消刚才的修改？", vbInformation + vbYesNo, "请确认")
If ii = vbYes Then
wbDN.adoRen.Recordset.CancelBatch
Set dtgRen.DataSource = adoRen
Call mod1.khRQing
End If

'wbDN.cmdRadd.Enabled = True
End Sub

Private Sub cmdRadd_Click()
Dim tt As String
Dim oo As Integer
On Error Resume Next
'frmRen.Enabled = True



        tt = "select * from khren where rid=" & Val(lblRid.Caption)
        adoLxr.Close
        adoLxr.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        wbDN.adoLxr.Update "khdh", txtKhDm.Text '客户代号
        wbDN.adoLxr.Update "khMan", txtMan.Text '联系人
        wbDN.adoLxr.Update "khHk", txtHk.Text '户口
         '性别
         If wbDN.optMan.Value = True Then
            wbDN.lblXb.Caption = "男"
         ElseIf wbDN.optWoman.Value = True Then
            wbDN.lblXb.Caption = "女"
         End If
        wbDN.adoLxr.Update "khSex", wbDN.lblXb.Caption
        wbDN.adoLxr.Update "khSr", DateSerial(Year(dtpSr.Value), Month(dtpSr.Value), Day(dtpSr.Value)) '生日
        wbDN.adoLxr.Update "khZw", wbDN.txtZw.Text '职务
        wbDN.adoLxr.Update "khDpho", wbDN.txtLpho.Text '电话
        wbDN.adoLxr.Update "khDwadr", wbDN.txtLdwdz.Text '单位地址
        wbDN.adoLxr.Update "khJpho", wbDN.txtLjpho.Text '家庭电话
        wbDN.adoLxr.Update "khMob", wbDN.txtLjmob.Text '手机
        wbDN.adoLxr.Update "khJadr", wbDN.txtLjadr.Text '家庭地址
        For oo = 0 To 82
            wbDN.adoLxr.Update "kh" & oo, wbDN.txtK(oo).Text
        Next
        wbDN.adoLxr.Update "lc", 1
        wbDN.adoLxr.Update "lcRen", mod1.DName
        wbDN.adoLxr.Update "lcUid", mod1.DHid
        wbDN.adoLxr.Update "kid", Val(wbDN.lblKid.Caption)
        wbDN.adoLxr.UpdateBatch
    cmdLeft.Enabled = True
    cmdRight.Enabled = True
    '更新联系人列表
'    tt = "vkhren1(" & Val(wbDN.lblKid.Caption) & ",'" & wbDN.lblYwy.Caption & "'," & Val(wbDN.lblXid.Caption) & ")"
'    wbDN.adoRen.Recordset.Close
'    wbDN.adoRen.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    wbDN.adoRen.Recordset.Requery
    Set wbDN.dtgRen.DataSource = wbDN.adoRen
    cmdNew.Enabled = True
    dtgRen.Visible = True
    cmdRdel.Enabled = True
    
'    If lblFwid.Caption = "" Then
'        '添加事务
'        Call mod1.EnventAdd("项目资料", txtMan.Text, lblLcRen.Caption, lblLcUid.Caption, lblRid.Caption, lblQM(0).Caption, "", "", mod1.DName, mod1.DHid, 0, lblRid.Caption)
'    End If
    lblLcRen.Caption = mod1.DName
    lblLcUid.Caption = mod1.DHid

End Sub

Private Sub cmdRdel_Click()
Dim tt As String
Dim oo As Integer
Dim ii As Integer
On Error Resume Next

If wbDN.cmdQm(0).Caption <> "" Then
    MsgBox "已经签过字,您无权删除!"
    Exit Sub
End If
ii = MsgBox("确定要删除此记录？", vbInformation + vbYesNo, "请确认")
If ii = vbYes Then
    'If wbDN.adoRen.Recordset.RecordCount > 1 Then
        wbDN.txtMan.Text = "" '联系人
        wbDN.txtHk.Text = "" '户口
        wbDN.optMan.Value = False '性别
        wbDN.optWoman.Value = False
        wbDN.dtpSr.Value = mod1.HMDa '生日
        wbDN.txtZw.Text = "" '职务
        wbDN.txtLpho.Text = "" '电话
        wbDN.txtLdwdz.Text = "" '单位地址
        wbDN.txtLjpho.Text = "" '家庭电话
        wbDN.txtLjmob.Text = "" '手机
        wbDN.txtLjadr.Text = "" '家庭地址


        For oo = 0 To 82
            wbDN.txtK(oo).Text = ""
        Next
        tt = "delete from khren where rid=" & Val(lblRid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        Set wbDN.dtgRen.DataSource = wbDN.adoRen
        '删除流程签字表中的记录
        tt = "delete from qmrz where qdbh='" & lblRid.Caption & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        '删除事务表中的记录
        tt = "delete from NewFuWu where fwid=" & Val(lblFwid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        If Dialog.Visible = True Then '更新事务列表
            Call mod1.refEnvent(1)
        End If
                
        '更新联系人列表
'        tt = "vkhren1(" & Val(wbDN.lblKid.Caption) & ",'" & wbDN.lblYwy.Caption & "')"
'        wbDN.adoRen.Recordset.Close
'        wbDN.adoRen.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        wbDN.adoRen.Recordset.Requery
        Set wbDN.dtgRen.DataSource = wbDN.adoRen
        wbDN.cmdNew.Enabled = True


    
 


    'Call mod1.khRQing
    adoRen.Recordset.MovePrevious
    If adoRen.Recordset.RecordCount = 0 Then
        wbDN.cmdRdel.Enabled = False
        wbDN.cmdRadd.Enabled = False
        wbDN.cmdNew.Enabled = True
        wbDN.tabRen.Enabled = False
    End If
''    Set dtgRen.DataSource = adoRen
'    dtgRen_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
End If

End Sub

Private Sub cmdRight_Click()
Dim tt As String
Dim oo As Integer
On Error Resume Next
    frmWait.Show
    frmWait.ZOrder 0
    frmWait.Refresh
    frmWait.faWait.Play
    adoRen.Recordset.MoveNext
    cmdLeft.Enabled = True
'tt = "vkhren2(" & wbDN.adoRen.Recordset.Fields("rid").Value & ")"
'wbDN.adoLxr.Close
'wbDN.adoLxr.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
'
'wbDN.txtMan.Text = wbDN.adoLxr.Fields("khMan").Value '联系人
'wbDN.txtHk.Text = wbDN.adoLxr.Fields("khHk").Value '户口
'
' wbDN.lblXb.Caption = wbDN.adoLxr.Fields("khSex").Value '性别
' If wbDN.lblXb.Caption = "男" Then
'    wbDN.optMan.Value = True
' ElseIf wbDN.lblXb.Caption = "女" Then
'    wbDN.optWoman.Value = True
' End If
'wbDN.dtpSr.Value = wbDN.adoLxr.Fields("khSr").Value '生日
'wbDN.txtZw.Text = wbDN.adoLxr.Fields("khZw").Value '职务
'wbDN.txtLpho.Text = wbDN.adoLxr.Fields("khDpho").Value '电话
'wbDN.txtLdwdz.Text = wbDN.adoLxr.Fields("khDwadr").Value '单位地址
'wbDN.txtLjpho.Text = wbDN.adoLxr.Fields("khJpho").Value '家庭电话
'wbDN.txtLjmob.Text = wbDN.adoLxr.Fields("khMob").Value '手机
'wbDN.txtLjadr.Text = wbDN.adoLxr.Fields("khJadr").Value '家庭地址
'
'For oo = 0 To 82
'    wbDN.txtK(oo).Text = wbDN.adoLxr.Fields("kh" & oo).Value
'Next
'wbDN.lblYwy.Caption = wbDN.adoLxr.Fields("ywy").Value
'wbDN.lblUid.Caption = wbDN.adoLxr.Fields("uid").Value
'wbDN.lblXywy.Caption = wbDN.adoLxr.Fields("xywy").Value
'wbDN.lblXuid.Caption = wbDN.adoLxr.Fields("xuid").Value
'wbDN.lblLc.Caption = wbDN.adoLxr.Fields("lc").Value
'wbDN.lblLcRen.Caption = wbDN.adoLxr.Fields("lcRen").Value
'wbDN.lblLcUid.Caption = wbDN.adoLxr.Fields("lcUid").Value
'wbDN.lblFwid.Caption = wbDN.adoLxr.Fields("Fwid").Value
'
''更新签字
'Call mod1.OpenKHAN
'frmWait.Visible = False

'adoRen.Recordset.MoveNext
'If adoRen.Recordset.EOF = True Then
'    cmdRight.Enabled = False
'End If
'adoRen.Recordset.MovePrevious
If adoRen.Recordset.EOF = True Then
    'cmdRight.Enabled = False
    adoRen.Recordset.MoveFirst
End If
End Sub

Private Sub cmdSave_Click()
Dim tt As String
On Error Resume Next
'If txtJc.Text = "" Then
'txtJc.Text = txtKhmc.Text
'End If
If tabKh.Tab = 1 And txtMan.Text <> "" Then
    Call cmdRadd_Click
End If

'检查是否有相同的项目名称.
tt = "select xmmc from xmzl where xmmc='" & txtXMMC.Text & "' and xid <>" & Val(lblXid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.HTP.RecordCount > 0 Then
    MsgBox ("您的项目名称有重复,请重新命名!")
    txtXMMC.SetFocus
    Exit Sub
End If
Call mod1.xmAdd
'如果项目平台小于60,而且为项目业务员本人的客户,则可以保存客户资料的修改.
If (lblYwy.Caption = lblKywy.Caption And lblUid.Caption = lblKuid.Caption And Val(lblXmPd.Caption) < 60) Or mod1.KhK = 1 Or lblKywy.Caption = "" Then
    Call mod1.khAdd '客户资料添加
End If


cmdLeft.Enabled = True
cmdRight.Enabled = True

dtgRen.Visible = True
If tabKh.TabEnabled(1) = False Then '如果为建立的客户
    tabKh.TabEnabled(1) = True
    cmdNew.Enabled = False
    cmdRdel.Enabled = False
End If
Call mod1.YJJL '意见交流
cmdMod.Enabled = True
optYz.Enabled = True
optWy.Enabled = True
frmLblQT.Enabled = True
cmdNew.Enabled = True
cmdSave.Enabled = False
dtgP.Visible = True
End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdSaveA_Click()
'If txtJc.Text = "" Then
'txtJc.Text = txtKhmc.Text
'End If
'If tabKh.TabEnabled(0) = False And tabKh.TabEnabled(2) = False Then '如果为合同评审单中保存，则只保存机组设备
''wbDN.adoA.Recordset.UpdateBatch
''wbDN.adoB.Recordset.UpdateBatch
''wbDN.adoC.Recordset.UpdateBatch
''wbDN.adoD.Recordset.UpdateBatch
''wbDN.adoE.Recordset.UpdateBatch
''wbDN.adoF.Recordset.UpdateBatch
''wbDN.adoG.Recordset.UpdateBatch
'Else '如果为客户资料中的保存，则保存全部客户信息
'Call mod1.khAdd '客户资料添加
'End If
'If wbDN.tabKh.TabEnabled(2) = False Then
'wbDN.adoRen.Recordset.AddNew "khdh", txtDm.Text
'wbDN.cmdRadd.Enabled = False
'End If
'wbDN.tabKh.TabEnabled(2) = True
''cmdDel.Enabled = False
'cmdSave.Enabled = False
''cmdSaveA.Enabled = False
End Sub



Private Sub Command1_Click()

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub comPz_Click()
Select Case comPz.ListIndex
Case 0
    txtPjPz.Text = "压缩机"
Case 1
    txtPjPz.Text = "水泵马达"
Case 2
    txtPjPz.Text = "风扇马达"
Case 3
    txtPjPz.Text = "风机马达"
Case Else
    txtPjPz.Text = ""
End Select
End Sub


Private Sub dtgB_DblClick()
'On Error Resume Next
'If adoB.Recordset.Fields("BF").Value = True Then
'adoB.Recordset.Fields("BF").Value = False
'ElseIf adoB.Recordset.Fields("BF").Value = False Then
'adoB.Recordset.Fields("BF").Value = True
'End If
'Set dtgB.DataSource = adoB
End Sub


Private Sub dtgB_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{tab}"
End Sub

'Private Sub dtGC_DblClick()
'On Error Resume Next
'If adoC.Recordset.Fields("BF").Value = True Then
'adoC.Recordset.Fields("BF").Value = False
'ElseIf adoC.Recordset.Fields("BF").Value = False Then
'adoC.Recordset.Fields("BF").Value = True
'End If
'Set dtGC.DataSource = adoC
'End Sub
'
'
'Private Sub dtGC_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub dtgD_DblClick()
'On Error Resume Next
'If adoD.Recordset.Fields("BF").Value = True Then
'adoD.Recordset.Fields("BF").Value = False
'ElseIf adoD.Recordset.Fields("BF").Value = False Then
'adoD.Recordset.Fields("BF").Value = True
'End If
'Set dtgD.DataSource = adoD
'End Sub
'
'
'Private Sub dtgD_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub dtgE_DblClick()
'On Error Resume Next
'If adoE.Recordset.Fields("BF").Value = True Then
'adoE.Recordset.Fields("BF").Value = False
'ElseIf adoE.Recordset.Fields("BF").Value = False Then
'adoE.Recordset.Fields("BF").Value = True
'End If
'Set dtgE.DataSource = adoE
'End Sub
'
'
'Private Sub dtgE_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub dtgF_DblClick()
'On Error Resume Next
'If adoF.Recordset.Fields("BF").Value = True Then
'adoF.Recordset.Fields("BF").Value = False
'ElseIf adoF.Recordset.Fields("BF").Value = False Then
'adoF.Recordset.Fields("BF").Value = True
'End If
'Set dtgF.DataSource = adoF
'End Sub


'Private Sub dtgF_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub dtgG_DblClick()
'On Error Resume Next
'If adoG.Recordset.Fields("BF").Value = True Then
'adoG.Recordset.Fields("BF").Value = False
'ElseIf adoG.Recordset.Fields("BF").Value = False Then
'adoG.Recordset.Fields("BF").Value = True
'End If
'Set dtgG.DataSource = adoG
'End Sub
'
'
'Private Sub dtgG_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{tab}"
'End Sub

Private Sub dtgA_Click()

End Sub

Private Sub dtgLouPan_Click()
Dim oo As Integer
On Error Resume Next
'adoRen.Recordset.MovePrevious
For oo = 0 To 4
    txtL(oo).Text = adoLouPan.Recordset.Fields("w" & oo).Value
    
Next
cmdGx.Enabled = True
End Sub

Private Sub dtgLouPan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim oo As Integer
On Error Resume Next
'adoRen.Recordset.MovePrevious
For oo = 0 To 4
    txtL(oo).Text = adoLouPan.Recordset.Fields("w" & oo).Value
    txtL(oo).Locked = False
Next
cmdGx.Enabled = True
End Sub


Private Sub dtgRen_Click()
frmQm.Visible = False
End Sub

Private Sub dtgRen_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim oo As Integer
Dim tt As String
On Error Resume Next
Call mod1.khRQing
    frmWait.Show
    frmWait.ZOrder 0
    frmWait.Refresh
    frmWait.faWait.Play
tt = "vkhren2(" & wbDN.adoRen.Recordset.Fields("rid").Value & ")"
wbDN.adoLxr.Close
wbDN.adoLxr.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
wbDN.txtMan.Text = wbDN.adoLxr.Fields("khMan").Value '联系人
wbDN.txtHk.Text = wbDN.adoLxr.Fields("khHk").Value '户口

 wbDN.lblXb.Caption = wbDN.adoLxr.Fields("khSex").Value '性别
 If wbDN.lblXb.Caption = "男" Then
    wbDN.optMan.Value = True
 ElseIf wbDN.lblXb.Caption = "女" Then
    wbDN.optWoman.Value = True
 End If
wbDN.dtpSr.Value = wbDN.adoLxr.Fields("khSr").Value '生日
wbDN.txtZw.Text = wbDN.adoLxr.Fields("khZw").Value '职务
wbDN.txtLpho.Text = wbDN.adoLxr.Fields("khDpho").Value '电话
wbDN.txtLdwdz.Text = wbDN.adoLxr.Fields("khDwadr").Value '单位地址
wbDN.txtLjpho.Text = wbDN.adoLxr.Fields("khJpho").Value '家庭电话
wbDN.txtLjmob.Text = wbDN.adoLxr.Fields("khMob").Value '手机
wbDN.txtLjadr.Text = wbDN.adoLxr.Fields("khJadr").Value '家庭地址
wbDN.lblRid.Caption = wbDN.adoLxr.Fields("rid").Value
For oo = 0 To 82
    wbDN.txtK(oo).Text = wbDN.adoLxr.Fields("kh" & oo).Value
Next
wbDN.lblYwy.Caption = wbDN.adoLxr.Fields("ywy").Value
wbDN.lblUid.Caption = wbDN.adoLxr.Fields("uid").Value
wbDN.lblXywy.Caption = wbDN.adoLxr.Fields("xywy").Value
wbDN.lblXuid.Caption = wbDN.adoLxr.Fields("xuid").Value
wbDN.lblLc.Caption = wbDN.adoLxr.Fields("lc").Value
wbDN.lblLcRen.Caption = wbDN.adoLxr.Fields("lcRen").Value
wbDN.lblLcUid.Caption = wbDN.adoLxr.Fields("lcUid").Value
wbDN.lblFwid.Caption = wbDN.adoLxr.Fields("Fwid").Value
'更新签字
Call mod1.OpenKHAN
frmWait.Visible = False
    cmdQm(0).Enabled = True
    cmdQm(1).Enabled = True
        wbDN.lblQM(0).Visible = False
    wbDN.lblQM(1).Visible = False
    wbDN.lblQM(2).Visible = False
    wbDN.lblTm(0).Visible = False
    wbDN.lblTm(1).Visible = False
    wbDN.lblTm(2).Visible = False
    wbDN.cmdQm(0).Visible = False
    wbDN.cmdQm(1).Visible = False
    wbDN.cmdQm(2).Visible = False
End Sub




Private Sub dtpBy_CloseUp()
txtK(7).Text = dtpBy.Value
txtK(7).SetFocus
End Sub

Private Sub dtpC_CloseUp()
'txtQrq.Text = dtpC.Value
'txtQrq.SetFocus
End Sub

Private Sub dtpJh_CloseUp()
txtK(17).Text = dtpJh.Value
txtK(17).SetFocus
End Sub




Private Sub dtpSr_CloseUp()
txtSr.Text = DateSerial(Year(dtpSr.Value), Month(dtpSr.Value), Day(dtpSr.Value))
End Sub

Private Sub Form_Click()
frmQm.Visible = False
End Sub

Public Sub QMBound(Rz, Lz As Integer)
Dim ii As Integer: Dim oo As Integer
On Error Resume Next
Call dtgPFF
dtgP.Rows = Lz + 20

For oo = 1 To Lz + 1
    dtgP.Row = oo
    For ii = 0 To 5
        dtgP.Col = ii
        dtgP.Text = Rz(ii, oo - 1)
        If ii = 3 Then
            If Len(Rz(ii, oo - 1)) > 16 Then
                dtgP.RowHeight(oo) = UpInt(Len(Rz(ii, oo - 1)) / 16) * dtgP.RowHeight(oo)
            End If
        End If
        If ii = 4 Then
            If dtgP.Text = "True" Then
                dtgP.Text = "同意"
            ElseIf dtgP.Text = "False" Then
                dtgP.Text = "驳回"
            End If

        End If
    Next
Next
For oo = 1 To Lz + 1
    dtgP.Row = oo
    dtgP.Col = 4
            If dtgP.Text = "驳回" Then
                For ii = 0 To 5
                    dtgP.Col = ii
                    dtgP.CellForeColor = &HFF&
                Next
            End If
Next
End Sub
Private Sub Form_Load()
Dim tt As String
Dim oo As Integer
On Error Resume Next
wbDN.Left = 0
wbDN.Top = 0
wbDN.Height = mod1.FHeight
wbDN.Width = mod1.FWidth
wbDN.Top = 0



frmLblQT.BorderStyle = 0

'设置区域
tt = "yzQyOpen"
Set adoQy = CreateObject("adodb.recordset")
Set adoLxr = CreateObject("adodb.recordset")
Set adoA = CreateObject("adodb.recordset")
Set adoKhmc = CreateObject("adodb.recordset")
tt = "select * from yzqy"
adoQy.Close
adoQy.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
Set comQy.RowSource = adoQy
comQy.ListField = "qy"
frmJz.BorderStyle = 0
frmJz.Left = 0
frmJz.Top = 1980
frmGL.Left = 0
frmGL.Top = 1980
wbDN.dtgA.ColWidth(0) = 300
wbDN.dtgA.ColWidth(1) = 1500
wbDN.dtgA.ColWidth(2) = 1700
wbDN.dtgA.ColWidth(3) = 3500
wbDN.dtgA.ColWidth(4) = 1000
wbDN.dtgA.ColWidth(5) = 1500
wbDN.dtgA.ColWidth(6) = 3500
wbDN.dtgA.ColWidth(7) = 1000
wbDN.dtgA.ColWidth(8) = 0
wbDN.dtgA.ColWidth(9) = 0
wbDN.dtgA.ColWidth(10) = 0
frmQm.Left = 0
frmQm.Top = 7480
'frmTT.Object
If mod1.ZT = "HBData" Then
    optWy.Visible = False
End If
End Sub

Private Sub Form_Resize()
'Call mod1.ResizeForm(Me) '确保窗体改变时控件随之改变


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim tt As String
'khAdd.Close
Dim ii As Integer
If MDI.Cq = False Then
If txtXMMC.Text = "" Then '判断项目名称
    ii = MsgBox("项目名称不能为空,退出将删除此项目!", vbInformation + vbYesNo, "询问")
    If ii = vbYes Then
        tt = "delete from xmzl where xid=" & Val(lblXid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Else
        Exit Sub
    End If
End If

If txtKhmc.Text = "" Then

        tt = "delete from khzl where kid=" & Val(lblKid.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

        '删除流程签字表中的记录
        tt = "delete from qmrz where qdbh='" & lblKid.Caption & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'End If
End If

wbDN.Visible = False
Cancel = True
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf frmKhBr.Visible = True Then
    frmKhBr.Show
    If wbDN.khAdd = True Then
        frmKhBr.tabCx.Tab = 0
        tt = "vkhNew('" & mod1.DName & "','" & mod1.DHid & "')"
        frmKhBr.adoKhBr.Close
        frmKhBr.adoKhBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        Set frmKhBr.dtgKh.DataSource = frmKhBr.adoKhBr
    End If
    frmKhBr.Enabled = True
    frmKhBr.ZOrder 0
ElseIf FMXC.Visible = True Then
    FMXC.Enabled = True
    FMXC.ZOrder 0
ElseIf FmxcNew.Visible = True Then
    FmxcNew.Enabled = True
    FmxcNew.ZOrder 0
End If
End If
frmHyxz.Visible = False
End Sub

Private Sub Label104_Click()

End Sub

Private Sub lblBm_Click()

End Sub

Private Sub lblQT_Click(Index As Integer)
Dim oo As Integer
Dim tt As String
Dim ii As Integer
On Error Resume Next

If wbDN.Visible = False Then Exit Sub
If wbDN.cmdSave.Enabled = True And txtKhmc.Text <> "" Then
    Call cmdSave_Click
End If
If lblQT(Index).Tag = 0 Then '如果为空,则添加项目资料
    Call mod1.khQing
    Call mod1.KhJQing
    Call mod1.khRQing
'    frmHyxz.Show
'    frmHyxz.ZOrder 0
'    wbDN.Enabled = False
'
'    dtgRen.Visible = True
    wbDN.tabKh.Tab = 0
    wbDN.tabKh.TabEnabled(1) = False
    Call mod1.XmKhUnLocked
Else '如果有客户名称,则显示此客户的资料
    Call mod1.khQing
    Call mod1.khRQing
    lblKid.Caption = lblQT(Index).Tag
    Call mod1.khBound(lblQT(Index).Tag, "qt" & Index)
    wbDN.tabKh.TabEnabled(1) = True
    'dtgRen.Visible = False
    If wbDN.adoRen.Recordset.RecordCount = 1 And wbDN.txtMan.Text = "" Then
        wbDN.cmdNew.Enabled = False
    Else
        wbDN.cmdNew.Enabled = True
    End If
    
    If wbDN.comXyxz.Text = "物业公司" Then
        frmGL.Visible = True
        frmJz.Visible = False
    Else
        frmGL.Visible = False
        frmJz.Visible = True
    End If
    
        If lblYwy.Caption = lblKywy.Caption Then
        '本人的客户, 能再编辑
        'wbDN.txtKhmc.Locked = True
        wbDN.comXz.Locked = False '企业性质
        wbDN.comXyxz.Locked = False '行业性质
        wbDN.txtAdr1.Locked = False '项目地址
        wbDN.comQy.Locked = False '区域
        wbDN.txtFH.Locked = False '国税号
        wbDN.txtKhYY.Locked = False '开户银行
        wbDN.txtZH.Locked = False '账号
        
        For oo = 0 To 4
            wbDN.txtL(oo).Locked = False
        Next
        
        wbDN.frmJE.Visible = True
        cmdSave.Enabled = False
        cmdMod.Enabled = True
    
    Else
        '别人的客户,不能再编辑
        'wbDN.txtKhmc.Locked = True
        wbDN.comXz.Locked = True '企业性质
        wbDN.comXyxz.Locked = True '行业性质
        wbDN.txtAdr1.Locked = True '项目地址
        wbDN.comQy.Locked = True '区域
        wbDN.txtFH.Locked = True '国税号
        wbDN.txtKhYY.Locked = True '开户银行
        wbDN.txtZH.Locked = True '账号
        
        For oo = 0 To 4
            wbDN.txtL(oo).Locked = True
        Next
        
        wbDN.frmJE.Visible = False
        cmdSave.Enabled = False
        cmdMod.Enabled = False
    End If
    
End If

End Sub

Private Sub optMan_Click()
lblXb.Caption = "男"
End Sub

Private Sub optQt_Click()
Dim ii As Integer
If wbDN.cmdSave.Enabled = True And txtKhmc.Text <> "" Then
    Call cmdSave_Click
End If
tabKh.Enabled = True
frmLblQT.Enabled = True
    Call mod1.khQing
    Call mod1.KhJQing
    Call mod1.khRQing
End Sub

Private Sub optWoman_Click()
lblXb.Caption = "女"
End Sub


Private Sub optWy_Click()
Dim tt As String
Dim oo As Integer
Dim ii As Integer
On Error Resume Next
If lblWy.Caption = "" Then Exit Sub
If wbDN.cmdSave.Enabled = True And txtKhmc.Text <> "" Then
    ii = MsgBox("切换客户单位前,是否保存所作的修改?", vbInformation + vbYesNo, "您好!")
    If ii = vbYes Then
        Call cmdSave_Click
    End If
End If
tabKh.Enabled = True
frmGL.Visible = True
frmLblQT.Enabled = False
For oo = 1 To 5
    lblQT(oo).Value = False
Next
wbDN.tabKh.Tab = 0
If lblWy.Tag = 0 Then '如果为空,则添加项目资料
    Call mod1.khQing
    Call mod1.KhJQing
    Call mod1.khRQing
    frmGL.Visible = True
    frmJz.Visible = False
    wbDN.tabKh.Tab = 0
    wbDN.tabKh.TabEnabled(1) = False
    Set wbDN.dtgLouPan.DataSource = Nothing
    'optYz.Enabled = False
    frmLblQT.Enabled = False
    Call mod1.XmKhUnLocked
Else '如果有客户名称,则显示此客户的资料
    Call mod1.khQing
    Call mod1.khRQing
    lblKid.Caption = lblWy.Tag
    Call mod1.khBound(lblWy.Tag, "wy")
    tabKh.TabEnabled(1) = True
        wbDN.comXyxz.Text = "物业公司"
    'dtgRen.Visible = False


    If lblYwy.Caption = lblKywy.Caption Then
        '本人的客户, 能再编辑
        'wbDN.txtKhmc.Locked = True
        wbDN.comXz.Locked = False '企业性质
        wbDN.comXyxz.Locked = False '行业性质
        wbDN.txtAdr1.Locked = False '项目地址
        wbDN.comQy.Locked = False '区域
        wbDN.txtFH.Locked = False '国税号
        wbDN.txtKhYY.Locked = False '开户银行
        wbDN.txtZH.Locked = False '账号
        
        For oo = 0 To 4
            wbDN.txtL(oo).Locked = False
        Next
        
        'wbDN.frmJE.Visible = True
        cmdSave.Enabled = False
        cmdMod.Enabled = True
    
    Else
        '别人的客户,不能再编辑
        'wbDN.txtKhmc.Locked = True
        wbDN.comXz.Locked = True '企业性质
        wbDN.comXyxz.Locked = True '行业性质
        wbDN.txtAdr1.Locked = True '项目地址
        wbDN.comQy.Locked = True '区域
        wbDN.txtFH.Locked = True '国税号
        wbDN.txtKhYY.Locked = True '开户银行
        wbDN.txtZH.Locked = True '账号
        
        For oo = 0 To 4
            wbDN.txtL(oo).Locked = True
        Next
        
        wbDN.frmJE.Visible = False
        cmdSave.Enabled = False
        cmdMod.Enabled = False
    End If
End If
    frmJz.Visible = False
    frmGL.Visible = True
End Sub

Private Sub optYz_Click()
Dim oo As Integer
Dim ii As Integer
Dim tt As String
On Error Resume Next
If wbDN.Visible = False Then Exit Sub
If wbDN.cmdSave.Enabled = True And txtKhmc.Text <> "" Then
    ii = MsgBox("切换客户单位前,是否保存所作的修改?", vbInformation + vbYesNo, "您好!")
    If ii = vbYes Then
        Call cmdSave_Click
    End If
End If
frmGL.Visible = False
frmJz.Visible = True
tabKh.TabEnabled(1) = True
frmLblQT.Enabled = False
For oo = 1 To 5
    lblQT(oo).Value = False
Next
If lblYz.Tag = 0 Then '如果为空,则添加项目资料
'    frmHyxz.Show
'    frmHyxz.ZOrder 0
'    wbDN.Enabled = False

    Call mod1.khQing
    Call mod1.khRQing
    Call mod1.KhJQing
   tabKh.Enabled = True
    wbDN.tabKh.Tab = 0
    wbDN.tabKh.TabEnabled(1) = False
    optWy.Enabled = False
    frmLblQT.Enabled = False
    Call mod1.XmKhUnLocked
Else '如果有客户名称,则显示此客户的资料
    Call mod1.khQing
    Call mod1.khRQing
    lblKid.Caption = lblYz.Tag
    Call mod1.khBound(lblYz.Tag, "yz")
    'dtgRen.Visible = False


    If lblYwy.Caption = lblKywy.Caption Then
        '本人的客户, 能再编辑
        'wbDN.txtKhmc.Locked = True
        wbDN.comXz.Locked = False '企业性质
        wbDN.comXyxz.Locked = False '行业性质
        wbDN.txtAdr1.Locked = False '项目地址
        wbDN.comQy.Locked = False '区域
        wbDN.txtFH.Locked = False '国税号
        wbDN.txtKhYY.Locked = False '开户银行
        wbDN.txtZH.Locked = False '账号
        
        For oo = 0 To 4
            wbDN.txtL(oo).Locked = False
        Next
        
        wbDN.frmJE.Visible = True
        cmdSave.Enabled = False
        cmdMod.Enabled = True
    
    Else
        '别人的客户,不能再编辑
        'wbDN.txtKhmc.Locked = True
        wbDN.comXz.Locked = True '企业性质
        wbDN.comXyxz.Locked = True '行业性质
        wbDN.txtAdr1.Locked = True '项目地址
        wbDN.comQy.Locked = True '区域
        wbDN.txtFH.Locked = True '国税号
        wbDN.txtKhYY.Locked = True '开户银行
        wbDN.txtZH.Locked = True '账号
        
        For oo = 0 To 4
            wbDN.txtL(oo).Locked = True
        Next
        
        wbDN.frmJE.Visible = False
        cmdSave.Enabled = False
        cmdMod.Enabled = False
    End If
End If
End Sub

Private Sub tabKh_Click(PreviousTab As Integer)
If tabKh.Tab = 1 Then
    'cmdSave.Enabled = True
    frmJz.Visible = False
    frmGL.Visible = False
   ' Call mod1.khRBound
Else
    If comXyxz.Text = "物业公司" Then
        frmJz.Visible = False
        frmGL.Visible = True
    Else
        frmJz.Visible = True
        frmGL.Visible = False
    End If
End If
End Sub

Private Sub txtDm_LostFocus()
'If txtKhmc.Text <> "" And txtDm.Text <> "" Then
'tabKh.TabEnabled(1) = True
'tabKh.TabEnabled(2) = True
'End If
End Sub


Private Sub tabRen_Click(PreviousTab As Integer)
frmQm.Visible = False
End Sub

Private Sub timQuit_Timer()
Dim oo As Integer
Dim ii As Integer
Dim Rz
Dim Lz As Integer
On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
Dim tt As String

If timZm = 1 Then '签字
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
    If Dialog.Visible = True Then
        Call mod1.refEnvent(1)
    End If
ElseIf timZm = 2 Then '新审核
    tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & lblXid.Caption & "' and yid=96 order by pid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rz = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lz = UBound(Rz, 2) + 1
    Call QMBound(Rz, Lz)
ElseIf timZm = 5 Then '联系人
    MsgBox "此联系人已经与合同关联!"
    If FmxcNew.Visible = True Then
        FmxcNew.txtYjBz.Text = FmxcNew.txtYjBz.Text & " " & LName
    End If
End If
timQuit.Enabled = False

End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
Dim LZw As String
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2,mt3 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    timWait.Enabled = False

    If timZm = 1 Then
        frmQm.Visible = False
        If OptT1.Value = True Then
            cmdQm(lblLc.Caption - 1).Caption = mod1.DName
            lblTm(lblLc.Caption - 1).Caption = mod1.DQda
        Else
            For oo = 0 To 5
                cmdQm(oo).Caption = ""
                lblTm(oo).Caption = ""
            Next
        End If
        lblLc.Caption = mod1.WP.Fields("mm1").Value
        lblFwid.Caption = mod1.WP.Fields("mm2").Value
        lblLcRen.Caption = mod1.WP.Fields("mt1").Value
        lblLcUid.Caption = mod1.WP.Fields("mt2").Value
        LZw = mod1.WP.Fields("mt3").Value
        
        If lblLc.Caption = "100" Then
            lblTX.Caption = "审核完毕！"
        Else
            lblTX.Caption = "下一流程,将跳至" & LZw & ": " & lblLcRen.Caption
        End If
    ElseIf timZm = 2 Then
                Lc = mod1.WP.Fields("mm1").Value
                Fwid = mod1.WP.Fields("mm2").Value
                LCRen = mod1.WP.Fields("mt1").Value
                LCUid = mod1.WP.Fields("mt2").Value
                lblTX.Caption = "下一流程,将跳至" & mod1.WP.Fields("mt3").Value & ": " & LCRen
                If Lc = 100 Then lblTX.Caption = "审核完毕!"
                
                



    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成
    
    
ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    timWait.Enabled = False
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    txtQM.Text = ""
    'lblRq.Caption = ""
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    txtQM.Text = ""
    'lblRq.Caption = ""
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub

Private Sub txtK_Change(Index As Integer)
If wbDN.Visible = False Then Exit Sub
If Len(txtK(Index).Text) > txtK(Index).Tag Then
MsgBox ("您编辑文字数将超过此项目的最大容纳字数,多余文字将不被保存!")
End If
End Sub

Private Sub txtK_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 7 Or Index = 17 Then
    If KeyCode = 8 Or KeyCode = 46 Then
        txtK(Index).Text = ""
    End If
End If
End Sub

Private Sub txtKhmc_Click(Area As Integer)
'Static Khmc As String '用来防止选择项目时,触发两次Click
'Dim oo As Integer
''If wbDN.Visible = False Or txtKhmc.Text = "" Or Khmc = txtKhmc.Text Then Exit Sub
'If wbDN.Visible = False Or txtKhmc.Text = "" Then Exit Sub
'
'If optYz.Value = True Then
'    Call mod1.khBound1(Val(txtKhmc.BoundText), "yz")
'    lblYz.Caption = txtKhmc.Text
'    lblYz.Tag = txtKhmc.BoundText
'ElseIf optWy.Value = True Then
'    Call mod1.khBound1(Val(txtKhmc.BoundText), "wy")
'    lblWy.Caption = txtKhmc.Text
'    lblWy.Tag = txtKhmc.BoundText
'ElseIf lblQT(1).Value = True Then
'    Call mod1.khBound1(Val(txtKhmc.BoundText), "qt1")
'    lblQT(1).Caption = txtKhmc.Text
'    lblQT(1).Tag = txtKhmc.BoundText
'ElseIf lblQT(2).Value = True Then
'    Call mod1.khBound1(Val(txtKhmc.BoundText), "qt2")
'    lblQT(2).Caption = txtKhmc.Text
'    lblQT(2).Tag = txtKhmc.BoundText
'ElseIf lblQT(3).Value = True Then
'    Call mod1.khBound1(Val(txtKhmc.BoundText), "qt3")
'    lblQT(3).Caption = txtKhmc.Text
'    lblQT(3).Tag = txtKhmc.BoundText
'ElseIf lblQT(4).Value = True Then
'    Call mod1.khBound1(Val(txtKhmc.BoundText), "qt4")
'    lblQT(4).Caption = txtKhmc.Text
'    lblQT(4).Tag = txtKhmc.BoundText
'ElseIf lblQT(5).Value = True Then
'    Call mod1.khBound1(Val(txtKhmc.BoundText), "qt5")
'    lblQT(5).Caption = txtKhmc.Text
'    lblQT(5).Tag = txtKhmc.BoundText
'End If
'Khmc = txtKhmc.Text
'
'If lblYwy.Caption <> lblKywy.Caption Then
'    '选择存在的客户,不能再编辑,是自己的客户除外.
'    'wbDN.txtKhmc.Locked = True
'    wbDN.comXz.Locked = True '企业性质
'    wbDN.comXyxz.Locked = True '行业性质
'    wbDN.txtAdr1.Locked = True '项目地址
'    wbDN.comQy.Locked = True '区域
'    wbDN.txtFH.Locked = True '国税号
'    wbDN.txtKhYY.Locked = True '开户银行
'    wbDN.txtZH.Locked = True '账号
'
'    wbDN.cmdLadd.Visible = False
'    wbDN.cmdLdel.Visible = False
'    wbDN.cmdGx.Visible = False
'    For oo = 0 To 4
'        wbDN.txtL(oo).Locked = True
'    Next
'
'    wbDN.frmJE.Visible = False
'Else
'    wbDN.frmJE.Visible = True
'End If



End Sub

Private Sub txtKhmc_KeyDown(KeyCode As Integer, Shift As Integer)
'Static Khmc As String
'Dim tt As String
'On Error Resume Next
'If txtKhmc.Text = "" Or wbDN.Visible = False Then Exit Sub
'If KeyCode = 13 Then
'    If optWy.Value = False Then
'        tt = "select khqc,kid from khzl where khqc like '%" & txtKhmc.Text & "%'"
'    Else
'        tt = "select khqc,kid from khzl where khqc like '%" & txtKhmc.Text & "%' and hyxz='物业公司'"
'    End If
'    Set adoKhmc = CreateObject("adodb.recordset")
'    adoKhmc.Close
'    adoKhmc.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    Set txtKhmc.RowSource = adoKhmc
'    txtKhmc.ListField = "khqc"
'    txtKhmc.BoundColumn = "kid"
'End If
End Sub

Private Sub txtKhmc_LostFocus()
Dim tt As String
Dim oo As Integer
On Error Resume Next
If Trim(txtKhmc.Text) = "" And Trim(lblKid.Caption) = "" And (lblKid.Caption <> "0") And optWy.Value = False Then
'If Trim(txtKhmc.Text) <> "" Then
    If optYz.Value = True Then
        lblYz.Caption = txtKhmc.Text
    ElseIf optWy.Value = True Then
        lblWy.Caption = txtKhmc.Text
    ElseIf lblQT(1).Value = True Then
        lblQT(1).Caption = txtKhmc.Text
    ElseIf lblQT(2).Value = True Then
        lblQT(2).Caption = txtKhmc.Text
    ElseIf lblQT(3).Value = True Then
        lblQT(3).Caption = txtKhmc.Text
    ElseIf lblQT(4).Value = True Then
        lblQT(4).Caption = txtKhmc.Text
    ElseIf lblQT(5).Value = True Then
        lblQT(5).Caption = txtKhmc.Text
    End If
ElseIf txtKhmc.Text <> "" And (lblKid.Caption = "" Or lblKid.Caption = "0") And optWy.Value = False Then '非物业客户添加
    frmHyxz.Show
    frmHyxz.ZOrder 0
    wbDN.Enabled = False
ElseIf txtKhmc.Text <> "" And (lblKid.Caption = "" Or lblKid.Caption = "0") And optWy.Value = True Then '物业客户添加
    '先取得代号编码


    tt = "Select max(khDh) as cou from khzl where khDh like '%WYG%'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    If IsNull(mod1.HTP.Fields("cou").Value) = True Then
    wbDN.txtKhDm.Text = "WYG" & Format(1, "0000")
    Else
    wbDN.txtKhDm.Text = "WYG" & Format(Val(Right(mod1.HTP.Fields("cou").Value, 4)) + 1, "0000")
    End If
    wbDN.tabKh.Enabled = True
    
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "khjia"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@Uid") = mod1.DHid
    mod1.cmd.Parameters("@khdh") = wbDN.txtKhDm.Text
    mod1.cmd.Parameters("@bm") = mod1.Bm
    mod1.cmd.Parameters("@hyxz") = "物业公司"
    mod1.cmd.Parameters("@Lcou") = Right(frmKhBr.cmdNew.ToolTipText, 1) '流程总数
    mod1.cmd.Parameters("@Lc") = 0 '当前流程
    mod1.cmd.Parameters("@lcRen") = mod1.DName
    mod1.cmd.Parameters("@lcUid") = mod1.DHid
    mod1.cmd.Parameters("@nLb") = frmKhBr.cmdNew.Tag
    mod1.cmd.Parameters("@xid") = Val(lblXid.Caption)
    mod1.cmd.Execute

    wbDN.comXyxz.Text = "物业公司"
    
    wbDN.lblKid.Caption = mod1.cmd.Parameters("@kid").Value
    wbDN.lblWy.Tag = mod1.cmd.Parameters("@kid").Value
    wbDN.lblRid.Caption = mod1.cmd.Parameters("@rid").Value
    wbDN.lblYwy.Caption = mod1.DName
    wbDN.lblUid.Caption = mod1.DHid
    wbDN.lblLcRen.Caption = mod1.DName
    wbDN.lblLcUid.Caption = mod1.DHid
    wbDN.lblXywy.Caption = mod1.DName
    wbDN.lblXuid.Caption = mod1.DHid

    Set cmd = Nothing
    tt = "Select * from khloPan where kid='" & wbDN.lblKid.Caption & "'"
    wbDN.adoLouPan.Recordset.Close
    wbDN.adoLouPan.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set wbDN.dtgLouPan.DataSource = wbDN.adoLouPan
    wbDN.cmdGx.Enabled = False
    wbDN.cmdLdel.Enabled = False
    wbDN.cmdLadd.Enabled = True
    For oo = 0 To 4
        wbDN.txtL(oo).Locked = True
    Next
    '设置流程按钮
    Call mod1.khLcBut(38)
    wbDN.tabKh.TabEnabled(1) = False
    
    '新添加的客户,可以编辑

    'wbDN.txtKhmc.Locked = True
    wbDN.comXz.Locked = False '企业性质
    wbDN.comXyxz.Locked = False '行业性质
    wbDN.txtAdr1.Locked = False '项目地址
    wbDN.comQy.Locked = False '区域
    wbDN.txtFH.Locked = False '国税号
    wbDN.txtKhYY.Locked = False '开户银行
    wbDN.txtZH.Locked = False '账号
    wbDN.cmdLadd.Visible = True
    wbDN.cmdLdel.Visible = True
    wbDN.cmdGx.Visible = True
    For oo = 0 To 4
        wbDN.txtL(oo).Locked = False
    Next
    lblWy.Caption = txtKhmc.Text
    'wbDN.frmJE.Visible = True
End If
End Sub

Private Sub txtMan_DblClick()
Dim tt As String
Dim HT As String
Dim ii As Integer
ii = MsgBox("是否确认将此联系人关联合同评审单?", vbYesNo + vbQuestion, "请确认")
If ii = vbNo Then Exit Sub
If txtLjmob.Text = "" Then
    MsgBox "请输入手机！"
    Exit Sub
End If
If txtZw.Text = "" Then
    MsgBox "请输入职务！"
    Exit Sub
End If
If FmxcNew.Visible = False Then
    HT = InputBox("请输入合同编号：")
Else
    HT = FmxcNew.txtHtbh.Text
End If
If HT = "" Then Exit Sub
timZm = 5 '联系人
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "新合同2011"
    mod1.cmd.Parameters("@NBLX") = "联系人"
    mod1.cmd.Parameters("@bh") = HT
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtMan.Text)
    mod1.cmd.Parameters("@mt2") = Trim(txtLjmob.Text)
    mod1.cmd.Parameters("@mlt1") = ""

    mod1.cmd.Parameters("@mm1") = 0

    LName = txtMan.Text & " " & txtLjmob.Text

            mod1.cmd.Parameters("@mb1") = 1


    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If
End Sub

Private Sub txtQrq_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Or KeyCode = 46 Then
    txtQrq.Text = ""
End If
End Sub


