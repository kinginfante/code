VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FMXC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������ʯ��ͬ����"
   ClientHeight    =   9180
   ClientLeft      =   -120
   ClientTop       =   330
   ClientWidth     =   15240
   ControlBox      =   0   'False
   Icon            =   "FMXC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin MSAdodcLib.Adodc adoFile 
      Height          =   375
      Left            =   10140
      Top             =   8010
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
         Name            =   "����"
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
      Left            =   9540
      Top             =   8070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FFFFC0&
      Caption         =   "������ǩ��ͬ"
      Height          =   285
      Left            =   -1410
      Style           =   1  'Graphical
      TabIndex        =   198
      Top             =   7770
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "������"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   -120
      TabIndex        =   81
      Top             =   6420
      Visible         =   0   'False
      Width           =   6315
      Begin VB.OptionButton optC 
         BackColor       =   &H00C0FFC0&
         Caption         =   "��ֹ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   5220
         TabIndex        =   249
         Top             =   1110
         Width           =   765
      End
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1305
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   85
         Top             =   300
         Width           =   4965
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ͬ��"
         Height          =   225
         Left            =   5220
         TabIndex        =   84
         Top             =   420
         Width           =   705
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�ܾ�"
         Height          =   195
         Left            =   5220
         TabIndex        =   83
         Top             =   780
         Width           =   675
      End
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "����"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame frmZt 
      Caption         =   "Frame3"
      Height          =   1125
      Left            =   10770
      TabIndex        =   112
      Top             =   8040
      Visible         =   0   'False
      Width           =   1545
      Begin VB.OptionButton optG 
         Caption         =   "�� �� ��"
         Height          =   195
         Left            =   210
         TabIndex        =   116
         Top             =   240
         Width           =   1035
      End
      Begin VB.OptionButton optP 
         Caption         =   "����׶�"
         Height          =   180
         Left            =   210
         TabIndex        =   115
         Top             =   930
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optZ 
         Caption         =   "ִ�н׶�"
         Height          =   225
         Left            =   210
         TabIndex        =   114
         Top             =   450
         Width           =   1035
      End
      Begin VB.OptionButton optW 
         Caption         =   "ִ�����"
         Height          =   225
         Left            =   210
         TabIndex        =   113
         Top             =   600
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdPje 
      Caption         =   "������"
      Height          =   1095
      Left            =   420
      TabIndex        =   96
      Top             =   8070
      Width           =   345
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   0
      Left            =   840
      TabIndex        =   95
      Top             =   8370
      Width           =   945
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   585
      Left            =   13950
      Picture         =   "FMXC.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   8580
      Width           =   645
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "�޸�"
      Height          =   585
      Left            =   12600
      Picture         =   "FMXC.frx":05CC
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   8580
      Width           =   645
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   585
      Left            =   14610
      Picture         =   "FMXC.frx":0A0E
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   8580
      Width           =   585
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "�ύ"
      Height          =   585
      Left            =   13260
      Picture         =   "FMXC.frx":0B10
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   8580
      Width           =   675
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8640
      Top             =   7860
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8220
      Top             =   8280
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   1
      Left            =   1860
      TabIndex        =   90
      Top             =   8370
      Width           =   945
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   2
      Left            =   2880
      TabIndex        =   89
      Top             =   8370
      Width           =   945
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   3
      Left            =   3870
      TabIndex        =   88
      Top             =   8370
      Width           =   945
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   4
      Left            =   4860
      TabIndex        =   87
      Top             =   8370
      Width           =   945
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   5
      Left            =   5880
      TabIndex        =   86
      Top             =   8370
      Width           =   945
   End
   Begin TabDlg.SSTab tabHt 
      Height          =   7665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   13520
      _Version        =   393216
      TabOrientation  =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "����"
      TabPicture(0)   =   "FMXC.frx":117A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frmDate"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmFk"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "comFP"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtgSD"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "comKQY"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdDZ"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtRGF2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdHt"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtZbh"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "optY2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "optY1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtYjpw"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "frmYm"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "timYj"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtTcRQ"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "frmYj"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtJlr2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtQt2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCbze2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtYf2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtFbje2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtRgf1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtCLF1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtFbje1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtYf1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtQt1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtClcb1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtCbze1"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtJlr1"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtClcb2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Frame2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtBz"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "frmFX"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtHtrq"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtZe"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtEd"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "comQy"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtXYwy"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtHtbh"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cmdWb"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtHtze"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtADR"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtKhdm"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtXMMC"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtKhmc"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtYwy"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "MMdtgFk"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Label9"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Label6"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "lblMF"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Label15"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Label5"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Label50"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Line1"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "lblRG"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "lblCL"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "lblFB"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "lblWC"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "lblCB"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "lblClcb"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "lblCBZE"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "lblJlr"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "Label49"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "lblHtxz"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "Label29"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "Label8"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "Label38"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "Label44"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "Label2(0)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "Label3(0)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "Label25"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "Label13"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "Label26"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "Label7"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "Label30"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "Line2"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "Line3"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).ControlCount=   78
      TabCaption(1)   =   "��������"
      TabPicture(1)   =   "FMXC.frx":1196
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "tabGc"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "��������"
      TabPicture(2)   =   "FMXC.frx":11B2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmCw"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame frmDate 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   -64170
         TabIndex        =   241
         Top             =   90
         Width           =   4125
         Begin VB.TextBox txtL 
            Height          =   300
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   243
            Top             =   270
            Width           =   1305
         End
         Begin VB.TextBox txtF 
            Height          =   300
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   242
            Top             =   270
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker dt4 
            Height          =   315
            Left            =   2400
            TabIndex        =   244
            Top             =   270
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��M��d��"
            Format          =   119275523
            CurrentDate     =   38098
         End
         Begin MSComCtl2.DTPicker dt3 
            Height          =   315
            Left            =   0
            TabIndex        =   245
            Top             =   270
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy��M��d��"
            Format          =   119275523
            CurrentDate     =   38098
         End
         Begin VB.Label Label28 
            Caption         =   "---��"
            Height          =   225
            Left            =   1920
            TabIndex        =   248
            Top             =   330
            Width           =   375
         End
         Begin VB.Label Label51 
            Caption         =   "ά����ʼ��"
            Height          =   225
            Left            =   210
            TabIndex        =   247
            Top             =   30
            Width           =   1605
         End
         Begin VB.Label Label52 
            Caption         =   "ά��������"
            Height          =   225
            Left            =   2520
            TabIndex        =   246
            Top             =   0
            Width           =   1275
         End
      End
      Begin VB.Frame frmFk 
         Height          =   915
         Left            =   -75000
         TabIndex        =   229
         Top             =   5370
         Width           =   4245
         Begin VB.OptionButton Option1 
            Caption         =   "���"
            Height          =   225
            Left            =   3120
            TabIndex        =   234
            Top             =   540
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton opt1 
            Caption         =   "���"
            Height          =   195
            Left            =   2220
            TabIndex        =   233
            Top             =   540
            Width           =   735
         End
         Begin VB.TextBox txtYje 
            Height          =   285
            Left            =   900
            TabIndex        =   232
            Top             =   480
            Width           =   1305
         End
         Begin VB.TextBox txtYrq 
            Height          =   300
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   231
            Top             =   150
            Width           =   1005
         End
         Begin VB.TextBox txtYed 
            Height          =   270
            Left            =   3150
            TabIndex        =   230
            Top             =   150
            Width           =   795
         End
         Begin MSComCtl2.DTPicker dtpYf 
            Height          =   315
            Left            =   900
            TabIndex        =   235
            Top             =   150
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   16711680
            CalendarTrailingForeColor=   8454016
            Format          =   119275521
            CurrentDate     =   38797
         End
         Begin VB.Label Label57 
            Caption         =   "�տ���"
            Height          =   225
            Left            =   60
            TabIndex        =   240
            Top             =   570
            Width           =   795
         End
         Begin VB.Label lblFid 
            Caption         =   "lblFid"
            Height          =   165
            Left            =   3600
            TabIndex        =   239
            Top             =   360
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label37 
            Caption         =   "%"
            Height          =   255
            Left            =   4050
            TabIndex        =   238
            Top             =   180
            Width           =   435
         End
         Begin VB.Label Label34 
            Caption         =   "�տ���"
            Height          =   255
            Left            =   2310
            TabIndex        =   237
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label33 
            Caption         =   "Ӧ������"
            Height          =   285
            Left            =   60
            TabIndex        =   236
            Top             =   180
            Width           =   735
         End
      End
      Begin VB.ComboBox comFP 
         Height          =   300
         ItemData        =   "FMXC.frx":11CE
         Left            =   -72510
         List            =   "FMXC.frx":11DB
         TabIndex        =   228
         Text            =   "Combo1"
         Top             =   3390
         Width           =   1545
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgSD 
         Height          =   2145
         Left            =   -75000
         TabIndex        =   226
         Top             =   5460
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   3784
         _Version        =   393216
         BackColor       =   12648384
         FixedCols       =   0
         BackColorFixed  =   16777152
         BackColorBkg    =   12648384
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.ComboBox comKQY 
         ForeColor       =   &H80000001&
         Height          =   300
         ItemData        =   "FMXC.frx":11FD
         Left            =   -68790
         List            =   "FMXC.frx":1213
         TabIndex        =   225
         Text            =   "Combo1"
         Top             =   1470
         Width           =   1305
      End
      Begin VB.CommandButton cmdDZ 
         Caption         =   "���ĵ��Ӻ�ͬ"
         Height          =   345
         Left            =   -71670
         TabIndex        =   223
         Top             =   1050
         Width           =   1305
      End
      Begin VB.TextBox txtRGF2 
         DataField       =   "UserName"
         DataSource      =   "adoFile"
         Height          =   270
         Left            =   -60870
         TabIndex        =   209
         Text            =   "Text2"
         Top             =   90
         Width           =   1185
      End
      Begin VB.CommandButton cmdHt 
         BackColor       =   &H008080FF&
         Caption         =   "BH"
         Height          =   225
         Left            =   -74130
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   405
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   2145
         Left            =   -74910
         MultiLine       =   -1  'True
         TabIndex        =   204
         Text            =   "FMXC.frx":123B
         Top             =   600
         Width           =   1065
      End
      Begin VB.TextBox txtZbh 
         Height          =   285
         Left            =   -73710
         Locked          =   -1  'True
         TabIndex        =   200
         Top             =   2370
         Width           =   1965
      End
      Begin VB.OptionButton optY2 
         Caption         =   "��"
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   -61860
         TabIndex        =   197
         Top             =   4290
         Width           =   885
      End
      Begin VB.OptionButton optY1 
         Caption         =   "����"
         ForeColor       =   &H00C000C0&
         Height          =   180
         Left            =   -63090
         TabIndex        =   196
         Top             =   4290
         Width           =   1035
      End
      Begin VB.TextBox txtYjpw 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -73710
         PasswordChar    =   "*"
         TabIndex        =   194
         Top             =   1890
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Frame frmCw 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   7125
         Left            =   -75000
         TabIndex        =   164
         Top             =   180
         Width           =   15225
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgyjF 
            Height          =   1935
            Left            =   5430
            TabIndex        =   193
            Top             =   120
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   3413
            _Version        =   393216
            Rows            =   10
            Cols            =   82
            BackColorBkg    =   16761024
            _NumberOfBands  =   1
            _Band(0).Cols   =   82
         End
         Begin VB.Frame frmQkF 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   1635
            Left            =   540
            TabIndex        =   184
            Top             =   450
            Width           =   4395
            Begin VB.TextBox txtQkfJe 
               Height          =   285
               Left            =   900
               TabIndex        =   188
               Top             =   600
               Width           =   2265
            End
            Begin VB.TextBox txtQkFBz 
               Height          =   555
               Left            =   900
               TabIndex        =   187
               Top             =   1020
               Width           =   2265
            End
            Begin VB.CommandButton cmdQkfAdd 
               Caption         =   "���"
               Height          =   285
               Left            =   3540
               TabIndex        =   186
               Top             =   1020
               Width           =   645
            End
            Begin VB.CommandButton cmdQkfDel 
               Caption         =   "ɾ��"
               Height          =   285
               Left            =   3510
               TabIndex        =   185
               Top             =   1320
               Width           =   675
            End
            Begin MSComCtl2.DTPicker dtpQkF 
               Height          =   285
               Left            =   900
               TabIndex        =   189
               Top             =   240
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   503
               _Version        =   393216
               Format          =   132841473
               CurrentDate     =   39312
            End
            Begin VB.Label Label48 
               Caption         =   "����"
               Height          =   225
               Left            =   150
               TabIndex        =   192
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label47 
               Caption         =   "���"
               Height          =   225
               Left            =   150
               TabIndex        =   191
               Top             =   630
               Width           =   615
            End
            Begin VB.Label Label46 
               Caption         =   "��ע"
               Height          =   225
               Left            =   150
               TabIndex        =   190
               Top             =   960
               Width           =   615
            End
         End
         Begin VB.CheckBox chkYJF 
            Caption         =   "����ҵ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   495
            Left            =   510
            TabIndex        =   182
            Top             =   0
            Width           =   1665
         End
         Begin VB.CheckBox chkJTF 
            Caption         =   "�ѽ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   465
            Left            =   510
            TabIndex        =   181
            Top             =   2310
            Width           =   1455
         End
         Begin VB.CheckBox chkQKF 
            Caption         =   "����ȫ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   435
            Left            =   510
            TabIndex        =   180
            Top             =   4710
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtYjf 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   179
            Top             =   120
            Width           =   2325
         End
         Begin VB.TextBox txtYjfBz 
            Height          =   795
            Left            =   360
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   178
            Top             =   5640
            Visible         =   0   'False
            Width           =   4545
         End
         Begin VB.TextBox txtJTf 
            Height          =   345
            Left            =   2400
            TabIndex        =   177
            Top             =   2370
            Width           =   2355
         End
         Begin VB.TextBox txtQkf 
            Height          =   285
            Left            =   2370
            TabIndex        =   175
            Top             =   4800
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Frame frmJTF 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   1665
            Left            =   540
            TabIndex        =   165
            Top             =   2820
            Width           =   4395
            Begin VB.TextBox txtJtfJe 
               Height          =   285
               Left            =   900
               TabIndex        =   169
               Top             =   600
               Width           =   2265
            End
            Begin VB.TextBox txtJTFbz 
               Height          =   555
               Left            =   900
               TabIndex        =   168
               Top             =   1020
               Width           =   2265
            End
            Begin VB.CommandButton cmdJTFadd 
               Caption         =   "���"
               Height          =   285
               Left            =   3540
               TabIndex        =   167
               Top             =   1020
               Width           =   645
            End
            Begin VB.CommandButton cmdJTFdel 
               Caption         =   "ɾ��"
               Height          =   285
               Left            =   3510
               TabIndex        =   166
               Top             =   1320
               Width           =   675
            End
            Begin MSComCtl2.DTPicker dtpJTF 
               Height          =   285
               Left            =   900
               TabIndex        =   170
               Top             =   240
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   503
               _Version        =   393216
               Format          =   132841473
               CurrentDate     =   39312
            End
            Begin VB.Label Label27 
               Caption         =   "����"
               Height          =   225
               Left            =   150
               TabIndex        =   173
               Top             =   300
               Width           =   735
            End
            Begin VB.Label Label40 
               Caption         =   "���"
               Height          =   225
               Left            =   150
               TabIndex        =   172
               Top             =   630
               Width           =   615
            End
            Begin VB.Label Label45 
               Caption         =   "��ע"
               Height          =   225
               Left            =   150
               TabIndex        =   171
               Top             =   960
               Width           =   615
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgQkf 
            Height          =   1755
            Left            =   5430
            TabIndex        =   174
            Top             =   4770
            Visible         =   0   'False
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   3096
            _Version        =   393216
            BackColorBkg    =   13172680
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgJTf 
            Height          =   2175
            Left            =   5430
            TabIndex        =   176
            Top             =   2280
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   3836
            _Version        =   393216
            Rows            =   8
            Cols            =   10
            BackColorBkg    =   12713983
            _NumberOfBands  =   1
            _Band(0).Cols   =   10
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            Index           =   0
            X1              =   0
            X2              =   15210
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00C000C0&
            BorderWidth     =   3
            Index           =   1
            X1              =   0
            X2              =   15210
            Y1              =   4620
            Y2              =   4620
         End
      End
      Begin VB.Frame frmYm 
         Caption         =   "��Ŀ������ϸ:"
         ForeColor       =   &H000000FF&
         Height          =   2265
         Left            =   -64560
         TabIndex        =   154
         Top             =   5400
         Width           =   4575
         Begin VB.CommandButton cmdYview 
            BackColor       =   &H00C0C0FF&
            Caption         =   "��ȯ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3990
            Style           =   1  'Graphical
            TabIndex        =   212
            Top             =   600
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "�ر�"
            Height          =   285
            Left            =   3990
            TabIndex        =   159
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtFED 
            Height          =   285
            Left            =   960
            TabIndex        =   158
            Top             =   1710
            Width           =   645
         End
         Begin VB.TextBox txtYingFu 
            Height          =   270
            Left            =   2880
            TabIndex        =   157
            Top             =   1710
            Width           =   1035
         End
         Begin VB.CommandButton cmdYadd 
            Caption         =   "���"
            Height          =   315
            Left            =   3990
            TabIndex        =   156
            Top             =   930
            Width           =   585
         End
         Begin VB.CommandButton cmdYdel 
            Caption         =   "ɾ��"
            Height          =   285
            Left            =   3990
            TabIndex        =   155
            Top             =   1290
            Width           =   585
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgYJ 
            Height          =   1275
            Left            =   150
            TabIndex        =   160
            Top             =   300
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   2249
            _Version        =   393216
            Rows            =   10
            Cols            =   6
            SelectionMode   =   1
            BorderStyle     =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
         Begin VB.Label lblyjFF 
            Caption         =   "lblYjff"
            Height          =   255
            Left            =   3180
            TabIndex        =   163
            Top             =   510
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label Label41 
            Caption         =   "�տ���"
            Height          =   255
            Left            =   120
            TabIndex        =   162
            Top             =   1740
            Width           =   825
         End
         Begin VB.Label Label39 
            Caption         =   "֧�����"
            Height          =   225
            Left            =   2010
            TabIndex        =   161
            Top             =   1740
            Width           =   915
         End
      End
      Begin VB.Timer timYj 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   -60180
         Top             =   840
      End
      Begin VB.TextBox txtTcRQ 
         Height          =   315
         Left            =   -61800
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "���ȡ������"
         Top             =   6060
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Frame frmYj 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   2775
         Left            =   -64530
         TabIndex        =   68
         Top             =   4500
         Visible         =   0   'False
         Width           =   4635
         Begin VB.ComboBox comYjRen 
            ForeColor       =   &H000000FF&
            Height          =   300
            ItemData        =   "FMXC.frx":127A
            Left            =   2610
            List            =   "FMXC.frx":127C
            TabIndex        =   210
            Text            =   "Combo1"
            Top             =   330
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton cmdCount 
            Caption         =   "����"
            Height          =   315
            Left            =   1800
            TabIndex        =   74
            Top             =   1740
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox txtTcBe 
            Height          =   285
            Left            =   1200
            TabIndex        =   73
            Text            =   "6"
            Top             =   1740
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txtTc2 
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   2100
            Width           =   1305
         End
         Begin VB.TextBox txtLr1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   1
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   71
            ToolTipText     =   "Ԥ��"
            Top             =   720
            Width           =   1185
         End
         Begin VB.TextBox txtYj1 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   330
            Width           =   1185
         End
         Begin VB.TextBox txtLr2 
            Height          =   285
            Left            =   2610
            Locked          =   -1  'True
            TabIndex        =   69
            ToolTipText     =   "ʵ��"
            Top             =   720
            Width           =   1215
         End
         Begin MSComCtl2.UpDown UpDa 
            Height          =   315
            Left            =   1530
            TabIndex        =   75
            Top             =   1740
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            Caption         =   "���˱���"
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   2640
            TabIndex        =   211
            ToolTipText     =   "˫���ɿ�������ϸ����"
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblTcBe 
            Caption         =   "��ɱ���"
            Height          =   195
            Left            =   270
            TabIndex        =   79
            Top             =   1800
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblLr 
            Caption         =   "�� �� 2"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   270
            TabIndex        =   78
            Top             =   780
            Width           =   915
         End
         Begin VB.Label lblTC 
            Caption         =   "��    ��"
            Height          =   195
            Left            =   270
            TabIndex        =   77
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label lblYj 
            Caption         =   "��Ŀ����"
            Height          =   255
            Left            =   270
            TabIndex        =   76
            Top             =   390
            Width           =   975
         End
      End
      Begin VB.TextBox txtJlr2 
         Height          =   285
         Left            =   -61920
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   3630
         Width           =   1215
      End
      Begin VB.TextBox txtQt2 
         Height          =   285
         Left            =   -60600
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   3210
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCbze2 
         Height          =   315
         Left            =   -61890
         Locked          =   -1  'True
         TabIndex        =   57
         ToolTipText     =   "ʵ��"
         Top             =   930
         Width           =   1185
      End
      Begin VB.TextBox txtYf2 
         Height          =   315
         Left            =   -62280
         Locked          =   -1  'True
         TabIndex        =   56
         ToolTipText     =   "ʵ��"
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtFbje2 
         Height          =   315
         Left            =   -61920
         Locked          =   -1  'True
         TabIndex        =   55
         ToolTipText     =   "ʵ��"
         Top             =   2730
         Width           =   1215
      End
      Begin VB.TextBox txtRgf1 
         Height          =   315
         Left            =   -63210
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1800
         Width           =   2475
      End
      Begin VB.TextBox txtCLF1 
         Height          =   285
         Left            =   -63210
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   2295
         Width           =   2505
      End
      Begin VB.TextBox txtFbje1 
         Height          =   285
         Left            =   -63210
         Locked          =   -1  'True
         TabIndex        =   52
         ToolTipText     =   "Ԥ��"
         Top             =   2730
         Width           =   1215
      End
      Begin VB.TextBox txtYf1 
         Height          =   285
         Left            =   -63930
         TabIndex        =   51
         ToolTipText     =   "Ԥ��"
         Top             =   210
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtQt1 
         Height          =   285
         Left            =   -63210
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   3180
         Width           =   2535
      End
      Begin VB.TextBox txtClcb1 
         Height          =   285
         Left            =   -63210
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1380
         Width           =   1215
      End
      Begin VB.TextBox txtCbze1 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   -63210
         Locked          =   -1  'True
         TabIndex        =   48
         ToolTipText     =   "Ԥ��"
         Top             =   930
         Width           =   1245
      End
      Begin VB.TextBox txtJlr1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   -63210
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   3630
         Width           =   1245
      End
      Begin VB.TextBox txtClcb2 
         Height          =   315
         Left            =   -61890
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1365
         Width           =   1185
      End
      Begin VB.Frame Frame2 
         Caption         =   "�ͻ�������:"
         ForeColor       =   &H000000FF&
         Height          =   3795
         Left            =   -70020
         TabIndex        =   36
         Top             =   3480
         Width           =   5235
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgFL 
            Height          =   3165
            Left            =   60
            TabIndex        =   221
            Top             =   300
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   5583
            _Version        =   393216
            Rows            =   10
            Cols            =   5
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorSel    =   0
            BackColorBkg    =   16777215
            SelectionMode   =   1
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   5
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "�ٴ���"
            Height          =   3315
            Left            =   2820
            TabIndex        =   213
            Top             =   270
            Width           =   1275
            Begin VB.TextBox txtD6 
               ForeColor       =   &H00008000&
               Height          =   270
               Left            =   120
               TabIndex        =   219
               Top             =   2520
               Width           =   1005
            End
            Begin VB.TextBox txtD5 
               ForeColor       =   &H00008000&
               Height          =   270
               Left            =   120
               TabIndex        =   218
               Top             =   2100
               Width           =   1005
            End
            Begin VB.TextBox txtD4 
               ForeColor       =   &H00008000&
               Height          =   270
               Left            =   120
               TabIndex        =   217
               Top             =   1680
               Width           =   1005
            End
            Begin VB.TextBox txtD3 
               ForeColor       =   &H00008000&
               Height          =   270
               Left            =   120
               TabIndex        =   216
               Top             =   1200
               Width           =   1005
            End
            Begin VB.TextBox txtD2 
               ForeColor       =   &H00008000&
               Height          =   270
               Left            =   120
               TabIndex        =   215
               Top             =   780
               Width           =   1005
            End
            Begin VB.TextBox txtD1 
               ForeColor       =   &H00008000&
               Height          =   270
               Left            =   120
               TabIndex        =   214
               Top             =   390
               Width           =   1005
            End
            Begin VB.Label Label4 
               Caption         =   "�ٴ���"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   225
               Left            =   180
               TabIndex        =   220
               Top             =   30
               Width           =   975
            End
         End
         Begin VB.TextBox txtFC 
            ForeColor       =   &H00004080&
            Height          =   285
            Left            =   1770
            TabIndex        =   202
            Top             =   3150
            Width           =   2325
         End
         Begin VB.TextBox txtH6 
            Height          =   270
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   2790
            Width           =   2325
         End
         Begin VB.TextBox txtH5 
            Height          =   270
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   2370
            Width           =   2325
         End
         Begin VB.TextBox txtH1 
            Height          =   270
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   660
            Width           =   2295
         End
         Begin VB.TextBox txtH2 
            Height          =   270
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   1050
            Width           =   2295
         End
         Begin VB.TextBox txtW3 
            Height          =   270
            Left            =   1770
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   1470
            Width           =   2325
         End
         Begin VB.TextBox txtW4 
            Height          =   270
            Left            =   1770
            TabIndex        =   41
            Top             =   1950
            Width           =   2295
         End
         Begin VB.TextBox txtW5 
            Height          =   270
            Left            =   1770
            TabIndex        =   40
            Top             =   2370
            Width           =   1035
         End
         Begin VB.TextBox txtW6 
            Height          =   270
            Left            =   1770
            TabIndex        =   38
            Top             =   2790
            Width           =   1035
         End
         Begin VB.Label lblYug 
            Caption         =   "Ԥ���ɱ�"
            Height          =   195
            Left            =   1860
            TabIndex        =   203
            Top             =   300
            Width           =   765
         End
         Begin VB.Label Label18 
            Caption         =   "����"
            ForeColor       =   &H00004080&
            Height          =   225
            Left            =   870
            TabIndex        =   201
            Top             =   3210
            Width           =   615
         End
         Begin VB.Label chkF 
            Caption         =   "���Ϸ�(��Ʒ)"
            Height          =   255
            Left            =   210
            TabIndex        =   153
            Top             =   2820
            Width           =   1275
         End
         Begin VB.Label chkE 
            Caption         =   "���Ϸ�(���)"
            Height          =   255
            Left            =   210
            TabIndex        =   152
            Top             =   2394
            Width           =   1455
         End
         Begin VB.Label chkD 
            Caption         =   "�˹���(ˮ����)"
            Height          =   255
            Left            =   210
            TabIndex        =   151
            Top             =   1968
            Width           =   1425
         End
         Begin VB.Label chkC 
            Caption         =   "�˹���(���̷ְ�)"
            Height          =   255
            Left            =   210
            TabIndex        =   150
            Top             =   1542
            Width           =   1485
         End
         Begin VB.Label chkB 
            Caption         =   "�˹���(����)"
            Height          =   255
            Left            =   210
            TabIndex        =   149
            Top             =   1140
            Width           =   1485
         End
         Begin VB.Label chkA 
            Caption         =   "�˹���(ά��)"
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   210
            TabIndex        =   148
            Top             =   690
            Width           =   1335
         End
         Begin VB.Label lblYug2 
            Caption         =   "�˼۳ɱ�"
            Height          =   225
            Left            =   2940
            TabIndex        =   45
            Top             =   300
            Width           =   915
         End
      End
      Begin VB.TextBox txtBz 
         Height          =   465
         Left            =   -68790
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   2460
         Width           =   3525
      End
      Begin VB.Frame frmFX 
         BorderStyle     =   0  'None
         Height          =   1605
         Left            =   -70830
         TabIndex        =   16
         Top             =   3600
         Width           =   585
         Begin VB.CommandButton cmdDe 
            Caption         =   "ɾ��"
            Height          =   375
            Left            =   0
            TabIndex        =   20
            Top             =   840
            Width           =   525
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "���"
            Height          =   375
            Left            =   0
            TabIndex        =   19
            Top             =   450
            Width           =   525
         End
         Begin VB.CommandButton cmdQing 
            Caption         =   "���"
            Height          =   345
            Left            =   0
            TabIndex        =   18
            Top             =   120
            Width           =   525
         End
         Begin VB.CommandButton cmdGx 
            Caption         =   "����"
            Height          =   315
            Left            =   0
            TabIndex        =   17
            Top             =   1230
            Visible         =   0   'False
            Width           =   525
         End
      End
      Begin VB.TextBox txtHtrq 
         Height          =   285
         Left            =   -73710
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   150
         Width           =   3315
      End
      Begin VB.TextBox txtZe 
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   -68790
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2970
         Width           =   1515
      End
      Begin VB.TextBox txtEd 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   -66120
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2970
         Width           =   885
      End
      Begin VB.ComboBox comQy 
         Height          =   300
         ItemData        =   "FMXC.frx":127E
         Left            =   -66180
         List            =   "FMXC.frx":1280
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1455
         Width           =   945
      End
      Begin VB.TextBox txtXYwy 
         Height          =   270
         Left            =   -68790
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1005
         Width           =   1245
      End
      Begin VB.TextBox txtHtbh 
         Height          =   270
         Left            =   -73710
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1620
         Width           =   3315
      End
      Begin VB.CommandButton cmdWb 
         BackColor       =   &H00008000&
         Caption         =   "��Ŀ����"
         Height          =   315
         Left            =   -71580
         TabIndex        =   9
         Top             =   2340
         Width           =   1185
      End
      Begin VB.TextBox txtHtze 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73710
         TabIndex        =   8
         ToolTipText     =   "���ڸ�����ϸ��ȷ����ͬ�ܽ��"
         Top             =   2760
         Width           =   3285
      End
      Begin VB.TextBox txtADR 
         Height          =   285
         Left            =   -68790
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2040
         Width           =   3555
      End
      Begin VB.TextBox txtKhdm 
         Height          =   270
         Left            =   -73710
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   1875
      End
      Begin VB.TextBox txtXMMC 
         Height          =   285
         Left            =   -68790
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   540
         Width           =   3555
      End
      Begin VB.ComboBox txtKhmc 
         Height          =   300
         Left            =   -73710
         Locked          =   -1  'True
         TabIndex        =   4
         ToolTipText     =   "�����б���ѡ��ͻ�"
         Top             =   570
         Width           =   3345
      End
      Begin VB.TextBox txtYwy 
         Height          =   270
         Left            =   -66540
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   990
         Width           =   1305
      End
      Begin TabDlg.SSTab tabGc 
         Height          =   7335
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   12938
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "ά��"
         TabPicture(0)   =   "FMXC.frx":1282
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label16"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label10"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label21"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label36"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label35"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label11"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "comZu"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "MMdtgA"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtWc"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtXc"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "frmTime"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtZu"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "����"
         TabPicture(1)   =   "FMXC.frx":129E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtDxnr"
         Tab(1).Control(1)=   "frmDx"
         Tab(1).Control(2)=   "txtZuD"
         Tab(1).Control(3)=   "MMdtgB"
         Tab(1).Control(4)=   "comZuD"
         Tab(1).Control(5)=   "Label14"
         Tab(1).Control(6)=   "Label12"
         Tab(1).Control(7)=   "Label56"
         Tab(1).Control(8)=   "Label55"
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "���"
         TabPicture(2)   =   "FMXC.frx":12BA
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "MMdtgMa"
         Tab(2).Control(1)=   "MMdtgBao"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "��Ʒ"
         TabPicture(3)   =   "FMXC.frx":12D6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "MMdtgCP"
         Tab(3).Control(1)=   "MMdtgCPCB"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "���̷ְ�"
         TabPicture(4)   =   "FMXC.frx":12F2
         Tab(4).ControlEnabled=   0   'False
         Tab(4).ControlCount=   0
         TabCaption(5)   =   "ˮ����"
         TabPicture(5)   =   "FMXC.frx":130E
         Tab(5).ControlEnabled=   0   'False
         Tab(5).ControlCount=   0
         Begin VB.TextBox txtDxnr 
            Height          =   3795
            Left            =   -74970
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   142
            Top             =   3000
            Width           =   15195
         End
         Begin VB.Frame frmDx 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   525
            Left            =   -67770
            TabIndex        =   138
            Top             =   1740
            Width           =   2865
            Begin VB.TextBox txtMon 
               Height          =   270
               Left            =   1260
               Locked          =   -1  'True
               TabIndex        =   139
               Top             =   210
               Width           =   525
            End
            Begin VB.Label Label22 
               Caption         =   "ά�ޱ�����"
               DragMode        =   1  'Automatic
               Height          =   225
               Left            =   120
               TabIndex        =   141
               Top             =   210
               Width           =   1065
            End
            Begin VB.Label Label23 
               Caption         =   "��"
               Height          =   255
               Left            =   1950
               TabIndex        =   140
               Top             =   210
               Width           =   195
            End
         End
         Begin VB.TextBox txtZuD 
            Height          =   285
            Left            =   -66540
            Locked          =   -1  'True
            TabIndex        =   131
            Text            =   "Text1"
            Top             =   945
            Width           =   1725
         End
         Begin VB.TextBox txtZu 
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   127
            Text            =   "Text1"
            Top             =   3195
            Width           =   1725
         End
         Begin VB.Frame frmTime 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   2025
            Left            =   11130
            TabIndex        =   123
            Top             =   270
            Width           =   3765
            Begin VB.CheckBox chkBc 
               Caption         =   "2Сʱ�ڵ���"
               Enabled         =   0   'False
               Height          =   255
               Left            =   450
               TabIndex        =   126
               Top             =   1290
               Width           =   1845
            End
            Begin VB.CheckBox chkBb 
               Caption         =   "ȫ����ת"
               Enabled         =   0   'False
               Height          =   255
               Left            =   450
               TabIndex        =   125
               Top             =   810
               Width           =   1845
            End
            Begin VB.CheckBox chkBa 
               Caption         =   "24Сʱ��ת"
               Enabled         =   0   'False
               Height          =   255
               Left            =   450
               TabIndex        =   124
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.TextBox txtXc 
            Height          =   270
            Left            =   10410
            Locked          =   -1  'True
            TabIndex        =   118
            Top             =   1560
            Width           =   405
         End
         Begin VB.TextBox txtWc 
            Height          =   270
            Left            =   8130
            Locked          =   -1  'True
            TabIndex        =   117
            Top             =   1560
            Width           =   495
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgA 
            Height          =   1545
            Left            =   90
            TabIndex        =   122
            Top             =   480
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   2725
            _Version        =   393216
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSDataListLib.DataCombo comZu 
            Height          =   330
            Left            =   1560
            TabIndex        =   128
            Top             =   2790
            Visible         =   0   'False
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   582
            _Version        =   393216
            Locked          =   -1  'True
            Text            =   "DataCombo2"
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgB 
            Height          =   1635
            Left            =   -74910
            TabIndex        =   132
            Top             =   600
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   2884
            _Version        =   393216
            SelectionMode   =   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin MSDataListLib.DataCombo comZuD 
            Height          =   330
            Left            =   -66540
            TabIndex        =   133
            Top             =   540
            Visible         =   0   'False
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   582
            _Version        =   393216
            Locked          =   -1  'True
            Text            =   "DataCombo2"
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgMa 
            Height          =   1155
            Left            =   -75000
            TabIndex        =   145
            Top             =   5640
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   2037
            _Version        =   393216
            BackColor       =   11927477
            Rows            =   5
            Cols            =   20
            BackColorBkg    =   -2147483627
            FillStyle       =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   20
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgCP 
            Height          =   5145
            Left            =   -75000
            TabIndex        =   146
            Top             =   30
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   9075
            _Version        =   393216
            Rows            =   30
            Cols            =   20
            BackColorBkg    =   -2147483627
            FillStyle       =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   20
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgCPCB 
            Height          =   1275
            Left            =   -75000
            TabIndex        =   147
            Top             =   5550
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   2249
            _Version        =   393216
            BackColor       =   11927477
            Rows            =   5
            Cols            =   20
            BackColorBkg    =   -2147483627
            FillStyle       =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   20
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgBao 
            Height          =   5175
            Left            =   -74940
            TabIndex        =   144
            Top             =   30
            Width           =   15225
            _ExtentX        =   26855
            _ExtentY        =   9128
            _Version        =   393216
            Rows            =   30
            Cols            =   20
            BackColorBkg    =   -2147483627
            FillStyle       =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            _NumberOfBands  =   1
            _Band(0).Cols   =   20
            _Band(0).GridLinesBand=   1
            _Band(0).TextStyleBand=   0
            _Band(0).TextStyleHeader=   0
         End
         Begin VB.Label Label14 
            Caption         =   "��������"
            Height          =   255
            Left            =   -74760
            TabIndex        =   143
            Top             =   2610
            Width           =   1785
         End
         Begin VB.Label Label12 
            Caption         =   "������Ϣ"
            Height          =   225
            Left            =   -74730
            TabIndex        =   137
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label11 
            Caption         =   "������Ϣ"
            Height          =   255
            Left            =   240
            TabIndex        =   136
            Top             =   180
            Width           =   1995
         End
         Begin VB.Label Label56 
            Caption         =   "���̲��鳤"
            Height          =   225
            Left            =   -67740
            TabIndex        =   135
            Top             =   1005
            Width           =   915
         End
         Begin VB.Label Label55 
            Caption         =   "���̲����"
            Height          =   225
            Left            =   -67740
            TabIndex        =   134
            Top             =   615
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label35 
            Caption         =   "���̲����"
            Height          =   225
            Left            =   270
            TabIndex        =   130
            Top             =   2865
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label36 
            Caption         =   "���̲��鳤"
            Height          =   225
            Left            =   270
            TabIndex        =   129
            Top             =   3255
            Width           =   945
         End
         Begin VB.Label Label21 
            Caption         =   "��"
            Height          =   225
            Left            =   10920
            TabIndex        =   121
            Top             =   1590
            Width           =   315
         End
         Begin VB.Label Label10 
            Caption         =   "�������"
            Height          =   225
            Left            =   9510
            TabIndex        =   120
            Top             =   1590
            Width           =   825
         End
         Begin VB.Label Label16 
            Caption         =   "ά������:"
            Height          =   225
            Left            =   7140
            TabIndex        =   119
            Top             =   1590
            Width           =   855
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgFk 
         Height          =   1665
         Left            =   -75000
         TabIndex        =   22
         Top             =   3720
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   2937
         _Version        =   393216
         Rows            =   50
         Cols            =   5
         FillStyle       =   1
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label9 
         Caption         =   "��Ʊ���ͣ�"
         Height          =   255
         Left            =   -73500
         TabIndex        =   227
         Top             =   3450
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "��������"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -69870
         TabIndex        =   224
         Top             =   1530
         Width           =   825
      End
      Begin VB.Label lblMF 
         Caption         =   "Label6"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74940
         TabIndex        =   222
         Top             =   3210
         Width           =   3255
      End
      Begin VB.Label Label15 
         Caption         =   "���ʽ"
         Height          =   195
         Left            =   -74910
         TabIndex        =   199
         Top             =   3450
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "��Ŀ����"
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   -64140
         TabIndex        =   195
         Top             =   4290
         Width           =   825
      End
      Begin VB.Label Label50 
         Caption         =   "%"
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   -65220
         TabIndex        =   183
         Top             =   3000
         Width           =   225
      End
      Begin VB.Line Line1 
         X1              =   -64800
         X2              =   -64800
         Y1              =   0
         Y2              =   7320
      End
      Begin VB.Label lblRG 
         Caption         =   "�˹�"
         Height          =   255
         Left            =   -63840
         TabIndex        =   67
         Top             =   1875
         Width           =   435
      End
      Begin VB.Label lblCL 
         Caption         =   "����"
         Height          =   255
         Left            =   -63840
         TabIndex        =   66
         Top             =   2325
         Width           =   465
      End
      Begin VB.Label lblFB 
         Caption         =   "�ְ�"
         Height          =   255
         Left            =   -63840
         TabIndex        =   65
         Top             =   2775
         Width           =   435
      End
      Begin VB.Label lblWC 
         Caption         =   "ά�ַ���"
         Height          =   255
         Left            =   -64200
         TabIndex        =   64
         Top             =   3240
         Width           =   825
      End
      Begin VB.Label lblCB 
         Caption         =   "�ɱ�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   -62940
         TabIndex        =   63
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblClcb 
         Caption         =   "����"
         Height          =   255
         Left            =   -63840
         TabIndex        =   62
         Top             =   1410
         Width           =   465
      End
      Begin VB.Label lblCBZE 
         Caption         =   "�ɱ��ܶ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64260
         TabIndex        =   61
         Top             =   960
         Width           =   885
      End
      Begin VB.Label lblJlr 
         Caption         =   "�� �� 1"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -64110
         TabIndex        =   60
         Top             =   3690
         Width           =   765
      End
      Begin VB.Label Label49 
         Caption         =   "��ע"
         Height          =   225
         Left            =   -69570
         TabIndex        =   35
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label lblHtxz 
         Height          =   315
         Left            =   -73680
         TabIndex        =   34
         Top             =   2040
         Width           =   3315
      End
      Begin VB.Label Label29 
         Caption         =   "�տ��ܶ�"
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   -69900
         TabIndex        =   33
         Top             =   3030
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "�տ���"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   -67050
         TabIndex        =   32
         Top             =   3000
         Width           =   915
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "��   ͬ   ��   ��   ��"
         BeginProperty Font 
            Name            =   "��Բ"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   345
         Left            =   -70140
         TabIndex        =   31
         Top             =   60
         Width           =   4485
      End
      Begin VB.Label Label44 
         Caption         =   "��  ��"
         Height          =   255
         Left            =   -66810
         TabIndex        =   30
         Top             =   1515
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "��Ŀ������"
         Height          =   255
         Index           =   0
         Left            =   -70050
         TabIndex        =   29
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label Label3 
         Caption         =   "��    ��"
         Height          =   255
         Index           =   0
         Left            =   -74910
         TabIndex        =   28
         Top             =   225
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "��ͬ���"
         Height          =   225
         Left            =   -67830
         TabIndex        =   27
         Top             =   7980
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "��ͬ�ܽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74970
         TabIndex        =   26
         Top             =   2790
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "��Ŀ��ַ"
         Height          =   255
         Left            =   -69900
         TabIndex        =   25
         Top             =   2070
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   "��Ŀ����"
         Height          =   255
         Left            =   -69900
         TabIndex        =   24
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label30 
         Caption         =   "ǩ����"
         Height          =   255
         Left            =   -67260
         TabIndex        =   23
         Top             =   1050
         Width           =   555
      End
      Begin VB.Line Line2 
         X1              =   -70110
         X2              =   -64800
         Y1              =   3450
         Y2              =   3450
      End
      Begin VB.Line Line3 
         X1              =   -70110
         X2              =   -70110
         Y1              =   3450
         Y2              =   7170
      End
   End
   Begin VB.Label lblFwid 
      Caption         =   "lblFwid"
      Height          =   255
      Left            =   1800
      TabIndex        =   208
      Top             =   0
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblLcUid 
      Caption         =   "lblLcUid"
      Height          =   285
      Left            =   480
      TabIndex        =   207
      Top             =   690
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblLcRen 
      Caption         =   "lblLcRen"
      Height          =   285
      Left            =   0
      TabIndex        =   206
      Top             =   240
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblLc 
      Caption         =   "lblLc"
      Height          =   315
      Left            =   1170
      TabIndex        =   205
      Top             =   360
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblMHid 
      Caption         =   "lblHid"
      Height          =   285
      Left            =   6930
      TabIndex        =   111
      Top             =   8040
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblMQM 
      Caption         =   "ҵ��Ա"
      Height          =   225
      Index           =   0
      Left            =   900
      TabIndex        =   110
      Top             =   8100
      Width           =   585
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   840
      TabIndex        =   109
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label lblJiLI 
      Caption         =   "���ٴΰ��ύ��ť,�Ա�ˢ������"
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   13050
      TabIndex        =   108
      Top             =   8160
      Width           =   1725
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   1860
      TabIndex        =   107
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label lblMQM 
      Caption         =   "���۾���"
      Height          =   225
      Index           =   1
      Left            =   1920
      TabIndex        =   106
      Top             =   8100
      Width           =   825
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   2
      Left            =   2880
      TabIndex        =   105
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label lblMQM 
      Caption         =   "�����ܼ�"
      Height          =   225
      Index           =   2
      Left            =   2940
      TabIndex        =   104
      Top             =   8100
      Width           =   825
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   3
      Left            =   3870
      TabIndex        =   103
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label lblMQM 
      Caption         =   "��ͬ����"
      Height          =   225
      Index           =   3
      Left            =   3960
      TabIndex        =   102
      Top             =   8100
      Width           =   825
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   4
      Left            =   4860
      TabIndex        =   101
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label lblMQM 
      Caption         =   "��ִͬ��"
      Height          =   225
      Index           =   4
      Left            =   4920
      TabIndex        =   100
      Top             =   8100
      Width           =   825
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   5
      Left            =   5880
      TabIndex        =   99
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label lblMQM 
      Caption         =   "���ȷ��"
      Height          =   225
      Index           =   5
      Left            =   5940
      TabIndex        =   98
      Top             =   8100
      Width           =   885
   End
   Begin VB.Label lblTX 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   6870
      TabIndex        =   97
      Top             =   8790
      Width           =   5475
   End
End
Attribute VB_Name = "FMXC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public adoOid As New ADODB.Recordset '����Old���ӵ�ADO
'Public adoBx as object '�ɹ���(���)
'Public adoGx as object '�ɱ���(���)
'Public adoBxCP as object '�ɹ���(��Ʒ)
'Public adoGxCP as object '�ɱ���(��Ʒ)
'Public adoFFk as object 'Ԥ�Ƹ���
'Public adoYj as object '�ʽ��



'Public adoA as object
'Public adoB as object

Dim timZm As Integer '�����ύ��,��timWaitִ�еĺ�������ID(2 �����ͬ 3�½�ѯ�۵�(���),6�½�ѯ�۵�(��Ʒ),
'10ǩ��11���ɺ�ͬ���12ɾ����ͬ13����༭15��ɱ༭16ȫ��༭��17��ǩ��ͬ,18�½�ά��ѯ�۵� 19 ִ��֪ͨ 20 �汾����)

Dim liD As Long
Dim LLid As Long
Dim LLXX As Boolean '(�½��˹�ѯ�ۣ��������ѯ�ۣ�

Dim Pw As String

Public FO As Single '���ʽѡ��

Public OldF As Boolean

Dim Rid(0 To 20) As Long '������ϵ�˵�ѡ���ID
Public NewF As Integer

Private Sub chkD_Click()
'''If chkC.Value = 1 Then
'''    tabHt.Tab = 1
'''    tabGc.TabVisible(5) = True
'''
'''End If
End Sub


Private Sub chkYJF_Click()
If chkYJF.Value = 1 Then
    txtYjf.Text = mod1.DQda
Else
    txtYjf.Text = ""
End If
End Sub

Private Sub cmdAdd_Click()
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim RL
Dim ul
On Error GoTo ERRch
'If cmdSave.Enabled = True Then
'    MsgBox "���ȱ��棡"
'    Exit Sub
'End If
'''''''Set mod1.cmd = createobject("adodb.command")
'''''''mod1.cmd.ActiveConnection = mod1.CC
'''''''mod1.cmd.CommandText = "htFkAdd"
'''''''mod1.cmd.CommandType = adCmdStoredProc
'''''''mod1.cmd.Parameters("@rq") = txtYrq.Text
'''''''mod1.cmd.Parameters("@yingfJe") = Round(Val(txtHtze.Text) * Val(txtYed.Text) / 100, 2)
'''''''If opt1.Value = True Then
'''''''    mod1.cmd.Parameters("@yingfJe") = Val(txtYje.Text)
'''''''End If
'''''''mod1.cmd.Parameters("@htbh") = lblMHid.Caption
'''''''mod1.cmd.Parameters("@ed") = Round(Val(txtYed.Text) / 100, 2)
'''''''If opt1.Value = True Then
'''''''    mod1.cmd.Parameters("@ed") = Round(Val(txtYje.Text) / Val(txtHtze.Text), 2)
'''''''End If
'''''''mod1.cmd.Execute
'''''''Set cmd = Nothing
'''''''
'''''''txtYed.Text = ""
'''''''mod1.mFk.Requery
'''''''Set MMdtgFk.DataSource = mod1.mFk
''''''''tt = "insert into htping1 (rq,yingfje,htbh,ed) values (@rq,@yingfje,@htbh,@ed)"
If opt1.Value = True Then
    tt = "insert into htping1 (rq,yingfje,htbh,ed) values ('" & DateSerial(Year(dtpYf.Value), Month(dtpYf.Value), Day(dtpYf.Value)) & "'," & Val(txtYje.Text) & _
            ",'" & lblMHid.Caption & "'," & Round(Val(txtYje.Text) / Val(txtHtze.Text), 2) & ")"
Else
    tt = "insert into htping1 (rq,yingfje,htbh,ed) values ('" & DateSerial(Year(dtpYf.Value), Month(dtpYf.Value), Day(dtpYf.Value)) & "'," & Round(Val(txtHtze.Text) * Val(txtYed.Text) / 100, 2) & _
            ",'" & lblMHid.Caption & "'," & Round(Val(txtYed.Text) / 100, 2) & ")"
End If
tt = tt & ";select Ӧ������,�տ���,Ӧ�����,fid from htFK where htbh='" & lblMHid.Caption & "' order by fid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = mod1.HTP.NextRecordset
RL = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next
ul = UBound(RL, 2) + 1

FMXC.MMdtgFk.Clear
FMXC.MMdtgFk.Rows = 30
FMXC.MMdtgFk.Row = 0: FMXC.MMdtgFk.Col = 1: FMXC.MMdtgFk.Text = "Ӧ������"
FMXC.MMdtgFk.Col = 2: FMXC.MMdtgFk.Text = "�տ���"
FMXC.MMdtgFk.Col = 3: FMXC.MMdtgFk.Text = "Ӧ�����"
For oo = 1 To ul + 1
    FMXC.MMdtgFk.Row = oo
    For ii = 1 To 4
        FMXC.MMdtgFk.Col = ii
        FMXC.MMdtgFk.Text = Trim(RL(ii - 1, oo - 1))
        If ii = 2 Then
            FMXC.MMdtgFk.Text = Str(Val(FMXC.MMdtgFk.Text) * 100) & "%"
        End If
    Next
Next
txtYed.Text = ""
Exit Sub
ERRch:
MsgBox ("������ϣ�������������Ϣ�����ԡ�")
End
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
If htBrow.Visible = True Then
'''    htBrow.adoBr.Requery
'''    Set htBrow.dtgBr.DataSource = htBrow.adoBr
    Call htBrow.dtgREF
    htBrow.Enabled = True
    htBrow.ZOrder 0
ElseIf FmxcXB.Visible = True Then
    FmxcXB.Enabled = True
    FmxcXB.ZOrder 0
ElseIf htBrowG.Visible = True Then
    htBrowG.Enabled = True
    htBrowG.ZOrder 0
ElseIf Dialog.Visible = True Then
    Call mod1.refEnvent(1)
    Dialog.ZOrder 0
    Dialog.Enabled = True
ElseIf frmGxBiao.Visible = True Then
    frmGxBiao.Enabled = True
    frmGxBiao.ZOrder 0
ElseIf frmCWBB.Visible = True Then
    frmCWBB.Enabled = True
    frmCWBB.ZOrder 0

End If
FmxcFK.Visible = False
End Sub

Private Sub cmdCGX_Click()
Dim CB As Long
Dim liD As Long
Dim Bid As Long
Dim XCB As Long
On Error Resume Next
If Val(txtCj.Text) = 0 Then Exit Sub
MMdtgBao.Col = 16
liD = MMdtgBao.Text
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "baoJiaGx"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@dj") = Val(txtCj.Text)
    mod1.cmd.Parameters("@sl") = Val(txtCL.Text)
    mod1.cmd.Parameters("@lid") = liD
    mod1.cmd.Execute
    'txtHg.Text = Val(txtHg.Text) + mod1.CMD.Parameters("@hg").Value
    Set cmd = Nothing
    
'    tt = "select bid from baojiaD where baoid=" & Val(lblBaoid.Caption)
'    Set mod1.HTP = CreateObject("adodb.recordset")
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    bid = mod1.HTP.Fields("bid").Value
'    If lblHtxz.Caption = "ά��" Or lblHtxz.Caption = "����" Then
'        '�����Ӧѯ�۵���cgid��
'        tt = "select cgid from xunJiaD where bid=" & bid
'        Set mod1.HTP = CreateObject("adodb.recordset")
'        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'        bid = mod1.HTP.Fields("cgid").Value
'    End If
'
'    '������Ӧѯ����ϸ�е�����
'    tt = "update XunJiaMx set sl=" & Val(txtTl.Text) & ",hg=dj*" & Val(txtTl.Text) & " where lid=" & liD
'    Set mod1.HTP = CreateObject("adodb.recordset")
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'    '������Ӧѯ�۵��еĽ��
'    tt = "select sum(hg) as hg from xunjiamx where bid=" & bid
'    Set mod1.HTP = CreateObject("adodb.recordset")
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''    XCB = 0
''    Do While Not mod1.HTP.EOF
''        XCB = XCB + mod1.HTP.Fields("hg").Value
''        mod1.HTP.MoveNext
''    Loop
'    XCB = mod1.HTP.Fields("hg").Value
'
'    tt = "update xunjiaD set hg=" & XCB & ",yhg=" & XCB & " where bid=" & bid
'    Set mod1.HTP = CreateObject("adodb.recordset")
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    txtCj.Text = ""
    txtCL.Text = ""
   ' txtClcb.Text = XCB
    mod1.mBxCP.Requery
    Set MMdtgCP.DataSource = mod1.mBxCP
   ' Call cmdSave_Click
    txtCj.Text = ""
    txtCL.Text = ""
End Sub

Private Sub cmdClose_Click()
frmYm.Visible = False
End Sub



Private Sub cmdD_Click()

End Sub

Private Sub cmdDe_Click()
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim RL
Dim ul
On Error Resume Next

tt = "delete from htfk where fid=" & Val(lblFid.Caption)
tt = tt & ";select Ӧ������,�տ���,Ӧ�����,fid from htFK where htbh='" & lblMHid.Caption & "' order by fid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = mod1.HTP.NextRecordset
RL = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
ul = UBound(RL, 2)
FMXC.MMdtgFk.Clear
FMXC.MMdtgFk.Rows = 30
FMXC.MMdtgFk.Row = 0: FMXC.MMdtgFk.Col = 1: FMXC.MMdtgFk.Text = "Ӧ������"
FMXC.MMdtgFk.Col = 2: FMXC.MMdtgFk.Text = "�տ���"
FMXC.MMdtgFk.Col = 3: FMXC.MMdtgFk.Text = "Ӧ�����"
For oo = 1 To ul + 1
    FMXC.MMdtgFk.Row = oo
    For ii = 1 To 4
        FMXC.MMdtgFk.Col = ii
        FMXC.MMdtgFk.Text = Trim(RL(ii - 1, oo - 1))
        If ii = 2 Then
            FMXC.MMdtgFk.Text = Str(Val(FMXC.MMdtgFk.Text) * 100) & "%"
        End If
    Next
Next
txtYed.Text = ""
End Sub



Private Sub cmdDel_Click()
Dim ii As Integer
If mod1.DName <> "������" And mod1.DName <> "�Ǽ���" And mod1.DName <> "�Ǽ���" And mod1.DName <> "�Ǽ���" Then
If Not (optZ.Value = False And (txtYwy.Text = mod1.DName Or txtXYwy.Text = mod1.DName)) Then Exit Sub
End If
ii = MsgBox("�Ƿ����ϴ˺�ͬ���󵥣�", vbYesNo + vbQuestion, "Hello")
If ii = vbNo Then
    Exit Sub
End If
timZm = 12 'ɾ����ͬ
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "ɾ����ͬ"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mm1") = 0
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
    mod1.cmd.Parameters("@mm19") = 0
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
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
Dim CJ As Double
Dim CJB As Double
Dim PP As Integer
Dim CM As String
On Error Resume Next
If OptT1.Value = True And lblMQM(Val(lblLc.Caption) - 1).Caption = "���ȷ��" And chkQKF.Value = 0 Then
    'If lblMQM(Index).Caption = "���ȷ��" Then
    If Val(txtZe.Text) < Val(txtHtze.Text) Then
            MsgBox ("δ��ȫ����ܵ���ɣ�")
            'frmQm.Visible = False
            Exit Sub
    End If
   ' End If
End If

If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("����һ��Ҫ���߾ܾ��ҵ�����!  :) ")
    Exit Sub
End If

If optC.Value = True And txtQM.Text = "" Then
    MsgBox ("����һ��Ҫд����ֹ������!  :) ")
    Exit Sub
End If
If optC.Value = False Then
    If OptT1.Value = True Then
        CM = "ͬ��"
    ElseIf optT2.Value = True Then
        CM = "����"
    Else
        Exit Sub
    End If
    
    PP = MsgBox("���Ƿ�ȷ�Ͻ���" & CM & "�Ĳ���?", vbYesNo + vbQuestion, "��������ȷ��ǩ�ֲ���!")
    If PP = vbNo Then Exit Sub
    
    frmFX.Visible = False
    timZm = 10 'ǩ��
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "��ͬ����"
        mod1.cmd.Parameters("@NBLX") = "ǩ��"
        mod1.cmd.Parameters("@bh") = Val(lblMHid.Caption)
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = txtYwy.Text
        mod1.cmd.Parameters("@mt2") = txtYwy.ToolTipText
        mod1.cmd.Parameters("@mt3") = txtXmmc.Text
        mod1.cmd.Parameters("@mt4") = txtHtbh.Text
        mod1.cmd.Parameters("@mt5") = ""
        mod1.cmd.Parameters("@mt6") = ""
        mod1.cmd.Parameters("@mt7") = ""
        mod1.cmd.Parameters("@mt8") = ""
        mod1.cmd.Parameters("@mt9") = ""
        mod1.cmd.Parameters("@mt10") = ""
        mod1.cmd.Parameters("@mt11") = ""
        mod1.cmd.Parameters("@mt12") = ""
        mod1.cmd.Parameters("@mt13") = ""
        mod1.cmd.Parameters("@mt14") = ""
        mod1.cmd.Parameters("@mt15") = lblHtxz.Caption
        mod1.cmd.Parameters("@mt16") = ""
        mod1.cmd.Parameters("@mt17") = ""
        mod1.cmd.Parameters("@mt18") = ""
        mod1.cmd.Parameters("@mt19") = ""
        mod1.cmd.Parameters("@mt20") = lblMQM(Val(lblLc.Caption) - 1).Caption
        mod1.cmd.Parameters("@mt21") = ""
        mod1.cmd.Parameters("@mt22") = ""
        mod1.cmd.Parameters("@mt23") = ""
        mod1.cmd.Parameters("@mt24") = ""
        mod1.cmd.Parameters("@mt25") = ""
        mod1.cmd.Parameters("@mlt1") = txtQM.Text '������
        mod1.cmd.Parameters("@mlt2") = ""
        mod1.cmd.Parameters("@mlt3") = ""
        mod1.cmd.Parameters("@mlt4") = ""
        mod1.cmd.Parameters("@mlt5") = ""
        mod1.cmd.Parameters("@mm1") = Val(lblLc.Caption)
        mod1.cmd.Parameters("@mm2") = Val(lblFwid.Caption)
        mod1.cmd.Parameters("@mm3") = 0
        mod1.cmd.Parameters("@mm4") = 0
        mod1.cmd.Parameters("@mm5") = 0
        mod1.cmd.Parameters("@mm6") = 0
        mod1.cmd.Parameters("@mm7") = 0
        mod1.cmd.Parameters("@mm8") = 0
        mod1.cmd.Parameters("@mm9") = 0
        mod1.cmd.Parameters("@mm10") = Val(txtHtze.Text)
        mod1.cmd.Parameters("@mm11") = Val(cmdW5.ToolTipText)
        mod1.cmd.Parameters("@mm12") = Val(cmdW6.ToolTipText)
        mod1.cmd.Parameters("@mm13") = 0
        mod1.cmd.Parameters("@mm14") = 1 '����֪ͨ
        mod1.cmd.Parameters("@mm15") = 0
        mod1.cmd.Parameters("@mm16") = 0
        mod1.cmd.Parameters("@mm17") = 0
        mod1.cmd.Parameters("@mm18") = 0
        mod1.cmd.Parameters("@mm19") = 0
        mod1.cmd.Parameters("@mm20") = 0
        If OptT1.Value = True Then
            mod1.cmd.Parameters("@mb1") = 1 'ͬ��
        Else
            mod1.cmd.Parameters("@mb1") = 0 '�ܾ�
        End If
    
        CJ = Val(txtHtze.Text) - Val(txtYj1.Text) - Val(txtCbze1.Text)
        'CJB = Abs(CJ) / Val(txtHtze.Text)
        CJB = CJ / Val(txtHtze.Text)
        If CJ < 0 Then
            'CJB = Abs(CJ) / Val(txtHtze.Text)
            mod1.cmd.Parameters("@mb2") = 0
        End If
        If Val(Right(lblMF.Caption, 4)) > 1 And Val(txtYj1.Text) = 0 Then
            mod1.cmd.Parameters("@mb2") = 1
        Else
            mod1.cmd.Parameters("@mb2") = 0
        End If
        If Val(lblLc.Caption) = 1 And optY1.Value = True Then
            mod1.cmd.Parameters("@mb2") = 0
        End If
        If optC.Value = True Then
            mod1.cmd.Parameters("@mb3") = 1
        Else
            mod1.cmd.Parameters("@mb3") = 0
        End If
        mod1.cmd.Parameters("@mb2") = 0
        mod1.cmd.Parameters("@mb4") = 0
        mod1.cmd.Parameters("@mb5") = 0
        mod1.cmd.Parameters("@md1") = Null
        mod1.cmd.Parameters("@md2") = Null
        mod1.cmd.Parameters("@md3") = Null
        mod1.cmd.Parameters("@md4") = Null
        mod1.cmd.Parameters("@md5") = Null
        Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
            MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
            cmdDing.Enabled = False
            Exit Sub
        Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
            Me.Enabled = False
            frmWaitA.Visible = True
            frmWaitA.Timer2.Enabled = False
            
            frmWaitA.ZOrder 0
            frmWaitA.Timer2.Enabled = True
            timWait.Enabled = True
        End If
            Set mod1.cmd = Nothing
ElseIf optC.Value = True Then
    CM = "��ֹ"
    
    PP = MsgBox("���Ƿ�ȷ�Ͻ���" & CM & "�Ĳ���?", vbYesNo + vbQuestion, "��������ȷ��ǩ�ֲ���!")
    If PP = vbNo Then Exit Sub
    
    frmFX.Visible = False
    timZm = 21 '��ֹ
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
        mod1.cmd.Parameters("@NBLX") = "��ֹ"
        mod1.cmd.Parameters("@bh") = Val(lblMHid.Caption)
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = txtYwy.Text
        mod1.cmd.Parameters("@mt2") = txtYwy.ToolTipText
        mod1.cmd.Parameters("@mt3") = txtXmmc.Text
        mod1.cmd.Parameters("@mlt1") = txtQM.Text '������
        mod1.cmd.Parameters("@mm1") = Val(lblLc.Caption)
        If OptT1.Value = True Then
            mod1.cmd.Parameters("@mb1") = 1 'ͬ��
        Else
            mod1.cmd.Parameters("@mb1") = 0 '�ܾ�
        End If
        mod1.cmd.Parameters("@md1") = Null
        Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
            MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
            cmdDing.Enabled = False
            Exit Sub
        Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
            Me.Enabled = False
            frmWaitA.Visible = True
            frmWaitA.Timer2.Enabled = False
            
            frmWaitA.ZOrder 0
            frmWaitA.Timer2.Enabled = True
            timWait.Enabled = True
        End If
            Set mod1.cmd = Nothing
End If
        

End Sub


Private Sub cmdGG_Click()

End Sub

Private Sub cmdDZ_Click()
If Val(txtHtbh.ToolTipText) = 0 Then Exit Sub
If mod1.DName <> txtYwy.Text And mod1.DName <> txtXYwy.Text And mod1.KhK = 0 And mod1.DName <> "�Ǽ���" And mod1.DName <> "������" And mod1.DName <> "���ĳ�" And mod1.DName <> "����" And mod1.Bm <> "����" And mod1.DName <> "����" And mod1.DName <> "����" Then Exit Sub

Dim bt() As Byte
Dim tt As String
On Error Resume Next

tt = "select fnr,fsize,fname from ht where fid=" & Val(txtHtbh.ToolTipText)
frmGGL.adoFile.Recordset.Close
frmGGL.adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockReadOnly, adCmdText
ReDim bt(frmGGL.adoFile.Recordset.Fields("Fsize").Value) As Byte
bt() = frmGGL.adoFile.Recordset.Fields("FNR").GetChunk(frmGGL.adoFile.Recordset.Fields("Fsize").Value + 1)

Open ("c:\work\demo\hmxp9000\" & frmGGL.adoFile.Recordset.Fields("fname").Value) For Binary As #2
Put #2, , bt()
Close #2

'tt = "Select * from hmfile where ywy='" & frmLogin.Combo1.Text & "'"
'frmFile.adoFile.Recordset.Close
'frmFile.adoFile.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'Set frmFile.dtGGF.DataSource = frmFile.adoFile
''�жϴ�����
'If adoFile.Recordset.Fields("Flx").Value = "WORD" Then
    frmGGL.OLE2.SourceDoc = "c:\work\demo\hmxp9000\" & frmGGL.adoFile.Recordset.Fields("fname").Value
    frmGGL.OLE2.Action = 1
    frmGGL.OLE2.DoVerb (-2)
    
'ElseIf adoFile.Recordset.Fields("Flx").Value = "EXCEL" Then
'    OLE1.SourceDoc = "c:\work\demo\file\" & FName
'    OLE1.Action = 1
'    OLE1.DoVerb (-2)
'End If
End Sub

Private Sub cmdGx_Click()
On Error Resume Next
'If cmdSave.Enabled = True Then
'    MsgBox "���ȱ��棡"
'    Exit Sub
'End If
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "htFkGx"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@rq") = dtpYf.Value
mod1.cmd.Parameters("@yingfJe") = Round(Val(txtHtze.Text) * Val(txtYed.Text) / 100, 2)
mod1.cmd.Parameters("@htbh") = Trim(lblMHid.Caption)
mod1.cmd.Parameters("@ed") = Round(Val(txtYed.Text) / 100, 2)
mod1.cmd.Parameters("@Fid") = Val(lblFid.Caption)
mod1.cmd.Execute
Set cmd = Nothing

txtYed.Text = ""
mod1.mFk.Requery
Set MMdtgFk.DataSource = mod1.mFk
End Sub

Private Sub cmdHt_Click()
Dim Ra, Rb, RC, RD, RE, Rf
Dim La, Lb, Lc, Ld, Le, LF
Dim Qy As String
Dim xZ As String
Dim XZDm As String
Dim tt As String
Dim ii As Integer
On Error Resume Next

Dim ZED As Double
Dim oo As Integer
Dim Zje As Double
Dim Tywy As String '������ת����һ�˵�����
Dim Tuid As String
Dim Oywy As String 'ԭ����ת�˵�����
Dim Ouid As String 'ԭ����ת�˵Ĺ���
Dim Bid1 As Long
Dim Bid6 As Long
Dim Bid7 As Long

'�ɰ治��������
If txtHtbh.Text = "HMNEW" And (lblLc.Caption = 1 Or lblLc.Caption = 0) And lblLcRen.Caption = mod1.DName Then
    dtgFL.Col = 4
    dtgFL.Row = 1: Bid1 = Val(Mid(dtgFL.Text, 4, Len(dtgFL.Text) - 3))
    dtgFL.Row = 6: Bid6 = Val(Mid(dtgFL.Text, 4, Len(dtgFL.Text) - 3))
    dtgFL.Row = 7: Bid7 = Val(Mid(dtgFL.Text, 4, Len(dtgFL.Text) - 3))
    If mod1.ZT = "HMData" Then
    oo = MsgBox("��Ϊ�ɰ��ͬ,��Ҫ������,������Ϊ�°�,�Ƿ�ͨ��������������?", vbInformation + vbYesNo, "����!")
        If oo = vbYes Then
            timZm = 20 '�汾����
                Set mod1.cmd = CreateObject("adodb.command")
                mod1.cmd.ActiveConnection = mod1.cc
                mod1.cmd.CommandText = "MLAdd"
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@zid") = 0
                mod1.cmd.Parameters("@errch") = ""
                mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
                mod1.cmd.Parameters("@NBLX") = "�汾����"
                mod1.cmd.Parameters("@bh") = Val(lblMHid.Caption)
                mod1.cmd.Parameters("@ywy") = mod1.DName
                mod1.cmd.Parameters("@uid") = mod1.DHid
                mod1.cmd.Parameters("@mt1") = lblHtxz.Caption
                mod1.cmd.Parameters("@mlt1") = ""
                mod1.cmd.Parameters("@mm1") = Bid1
                mod1.cmd.Parameters("@mm6") = Bid6
                mod1.cmd.Parameters("@mm7") = Bid7
                mod1.cmd.Parameters("@mb1") = 0
                Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
                mod1.cmd.Execute
                mod1.Zid = mod1.cmd.Parameters("@zid").Value
                If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
                    MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
                    cmdDing.Enabled = False
                    Exit Sub
                Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
                    Me.Enabled = False
                    frmWaitA.Visible = True
                    frmWaitA.Timer2.Enabled = False
                    
                    frmWaitA.ZOrder 0
                    frmWaitA.Timer2.Enabled = True
                    timWait.Enabled = True
                End If
            
                
            Set mod1.cmd = Nothing
    '''''    Else
    '''''        Exit Sub
            Exit Sub
        End If
            Exit Sub
    End If

'�жϺ�ͬ�еĸ�ѯ�۵�����ҵ��Աȷ��
tt = "select lc from xunjiaD where bid=" & Val(cmdW1.ToolTipText) & ";" & _
    "select lc from xunjiaD where bid=" & Val(cmdW5.ToolTipText) & ";" & _
    "select lc from xunjiaD where bid=" & Val(cmdW6.ToolTipText)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
La = 100: Lb = 100: Lc = 100: Ld = 100: Le = 100: LF = 100
If mod1.HTP.BOF = False Then
    Ra = mod1.HTP.GetRows
    La = Ra(0, 0)
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rb = mod1.HTP.GetRows
    Lb = Rb(0, 0)
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RC = mod1.HTP.GetRows
    Lc = RC(0, 0)
End If
mod1.HTP.Close
Set mod1.HTP = Nothing
If La <> 100 Then
    MsgBox "ά��ѯ�۵�û�гɱ�ȷ�ϣ�"
    Exit Sub
End If
If Lb <> 100 Then
    MsgBox "�����ѯ�۵�û�гɱ�ȷ�ϣ�"
    Exit Sub
End If
If Lc <> 100 Then
    MsgBox "��Ʒѯ�۵�û�гɱ�ȷ�ϣ�"
    Exit Sub
End If

dtgFL.Col = 2: dtgFL.Row = 5
'�жϺ�ͬ���ʺͺ�ͬ���.
If Val(txtH1.Text) > 0 Then
    ii = MsgBox("��ȷ�ϴ˵�������ǩ������ǩ��" & Chr(13) & Chr(10) & "�������ǡ�������ǩ�������񡱴�����ǩ��", vbYesNo + vbInformation, "����ȷ�ϣ�")
    xZ = "ά��"
    XZDm = "WB"
ElseIf Val(txtH2.Text) > 0 Then
    xZ = "����"
    XZDm = "DX"
ElseIf Val(txtW3.Text) > 0 Then
    xZ = "���̷ְ�"
    XZDm = "FB"
ElseIf Val(txtW4.Text) > 0 Then
    xZ = "ˮ����"
    XZDm = "WT"
ElseIf Val(dtgFL.Text) > 0 Then
    xZ = "��פ"
    XZDm = "CZ"
ElseIf Val(txtW5.Text) > 0 Or Val(txtH5.Text) > 0 Then
    xZ = "�����"
    XZDm = "LP"
ElseIf Val(txtW6.Text) > 0 Or Val(txtH6.Text) > 0 Then
    xZ = "��Ʒ"
    XZDm = "CP"
Else
    MsgBox "��ȷ���˿ͻ��������,�������ɺ�ͬ���!"
    Exit Sub
End If
'''''''dtgFL.Col = 2
'''''''dtgFL.Col = 1
'''''''If Val(dtgFL.Text) > 0 Then
'''''''    ii = MsgBox("��ȷ�ϴ˵�������ǩ������ǩ��" & Chr(13) & Chr(10) & "�������ǡ�������ǩ�������񡱴�����ǩ��", vbYesNo + vbInformation, "����ȷ�ϣ�")
'''''''    xZ = "ά��"
'''''''    XZDm = "WB"
'''''''ElseIf Val(txtH2.Text) > 0 Then
'''''''    xZ = "����"
'''''''    XZDm = "DX"
'''''''ElseIf Val(txtW3.Text) > 0 Then
'''''''    xZ = "���̷ְ�"
'''''''    XZDm = "FB"
'''''''ElseIf Val(txtW4.Text) > 0 Then
'''''''    xZ = "ˮ����"
'''''''    XZDm = "WT"
'''''''ElseIf Val(txtW5.Text) > 0 Or Val(txtH5.Text) > 0 Then
'''''''    xZ = "�����"
'''''''    XZDm = "LP"
'''''''ElseIf Val(txtW6.Text) > 0 Or Val(txtH6.Text) > 0 Then
'''''''    xZ = "��Ʒ"
'''''''    XZDm = "CP"
'''''''Else
'''''''    MsgBox "��ȷ���˿ͻ��������,�������ɺ�ͬ���!"
'''''''    Exit Sub
'''''''End If

If mod1.Qy = "�Ϻ�" Then
    Qy = "SH"
ElseIf mod1.Qy = "����" Then
    Qy = "HZ"
ElseIf mod1.Qy = "�Ͼ�" Then
    Qy = "NJ"
ElseIf mod1.Qy = "����" Then
    Qy = "BJ"
ElseIf mod1.Qy = "����" Then
    Qy = "GZ"
ElseIf mod1.Qy = "�人" Then
    Qy = "WH"
ElseIf mod1.Qy = "��̨" Then
    Qy = "YT"
ElseIf mod1.Qy = "֣��" Then
    Qy = "ZZ"
Else
    MsgBox "�µ�����,��δ�ں�����Ϣ��ע��,������������ϵ!"
    Exit Sub
End If
If mod1.ZT = "HMData" Then
    txtHtbh.Text = "HM" & Qy & Format(mod1.DQda, "yyyymmdd") & XZDm & lblMHid.Caption
Else
    txtHtbh.Text = "HB" & Qy & Format(mod1.DQda, "yyyymmdd") & XZDm & lblMHid.Caption
End If
    lblHtxz.Caption = xZ
    If xZ = "ά��" Then '��ͬ���ע����ǩ������ǩ
        If ii = vbYes Then
            txtHtbh.Text = "HN" & Qy & Format(mod1.DQda, "yyyymmdd") & XZDm & lblMHid.Caption
        Else
            txtHtbh.Text = "HO" & Qy & Format(mod1.DQda, "yyyymmdd") & XZDm & lblMHid.Caption
        End If
    End If
    
    timZm = 11 '���ɺ�ͬ���
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "��ͬ����"
        mod1.cmd.Parameters("@NBLX") = "��ͬ���"
        mod1.cmd.Parameters("@bh") = Val(lblMHid.Caption)
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = txtHtbh.Text
        mod1.cmd.Parameters("@mt2") = lblHtxz.Caption
        mod1.cmd.Parameters("@mt3") = ""
        mod1.cmd.Parameters("@mt4") = ""
        mod1.cmd.Parameters("@mt5") = ""
        mod1.cmd.Parameters("@mt6") = ""
        mod1.cmd.Parameters("@mt7") = ""
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
        mod1.cmd.Parameters("@mm1") = 0
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
        mod1.cmd.Parameters("@mm19") = 0
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
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
            MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
            cmdDing.Enabled = False
            Exit Sub
        Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
            Me.Enabled = False
            frmWaitA.Visible = True
            frmWaitA.Timer2.Enabled = False
    
            frmWaitA.ZOrder 0
            frmWaitA.Timer2.Enabled = True
            timWait.Enabled = True
        End If
End If
    
Set mod1.cmd = Nothing
cmdSave.Enabled = True
End Sub

Private Sub cmdJTFadd_Click()
Dim tt As String
Dim hg As Single
On Error Resume Next
If Val(txtJtfJe.Text) = 0 Then
Exit Sub
End If

timZm = 15 '��ӽ���
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "��ɱ༭"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text) '��ͬ���
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mlt1") = txtJTFbz.Text
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtJtfJe.Text)
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 1 '������
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = dtpJTF.Value
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


Set mod1.cmd = Nothing

End Sub

Private Sub cmdJTFdel_Click()
Dim tt As String
Dim ii As Integer
Dim Yid As Long

On Error Resume Next

dtgJTf.Col = 4
Yid = 0
Yid = Val(dtgJTf.Text)


If Yid = 0 Then
Exit Sub
End If



ii = MsgBox("�Ƿ�ɾ���˼�¼?", vbQuestion + vbYesNo, "ѯ��")
If ii = vbNo Then
    Exit Sub
End If

timZm = 15 '��ɱ༭
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "��ɱ༭"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text) '��ͬ���
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mm1") = Yid
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 0 '���ɾ��
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


Set mod1.cmd = Nothing
End Sub

Private Sub cmdMod_Click()
Dim oo As Integer

Dim Ra
Dim La
On Error Resume Next
For oo = 0 To 20
    Rid(oo) = 0
Next
If (mod1.DName = "�Ǽ���" Or mod1.DName = "�Ǽ���" Or mod1.DName = "�Ǽ���") And Val(lblLc.Caption) > 1 Then
'FMXC.txtJTf.Locked = False
'FMXC.txtQkf.Locked = False
'FMXC.chkYJF.Enabled = True
'FMXC.chkJTF.Enabled = True
'FMXC.chkQKF.Enabled = True
'FMXC.txtYjfBz.Locked = False
frmJTF.Visible = True
frmQkF.Visible = True
frmCw.Enabled = True
cmdSave.Enabled = True
cmdDel.Enabled = True
'''frmFk.Visible = True
    If mod1.DName = "�Ǽ���" Or mod1.DName = "�Ǽ���" Then
        cmdDel.Enabled = True
    End If

Exit Sub
End If
'''If lblLcUid.Caption <> mod1.DHid And Not (mod1.DName = "������") Then
'''Exit Sub
'''End If
cmdYadd.Visible = False
cmdYdel.Visible = False
txtYj1.Locked = True
comYjRen.Locked = True
If (txtXYwy.Text = mod1.DName Or txtYwy.Text = mod1.DName) Then

End If
If (lblLc.Caption = 1 Or lblLc.Caption = 0) And (txtXYwy.Text = mod1.DName Or txtYwy.Text = mod1.DName) Then
    frmFX.Visible = True
    dt3.Enabled = True
    dt4.Enabled = True
    optLa.Enabled = True
    optLb.Enabled = True
    optLc.Enabled = True
    cmdSave.Enabled = True
    txtHtze.Locked = False
    txtFbnr.Locked = False
    txtWBNR.Locked = False
    txtBz.Locked = False
    If mod1.Qy <> "�Ϻ�" And Val(lblMHid.Caption) < 19345 Then
        txtW3.Locked = False
        txtW4.Locked = False
    End If
    txtW5.Locked = False
    txtW6.Locked = False
    comKQY.Locked = False
    cmdW1.Visible = True: cmdW2.Visible = True: cmdW3.Visible = True: cmdW4.Visible = True: cmdW5.Visible = True: cmdW6.Visible = True
    frmFk.Visible = True
    comFP.Locked = False
    If mod1.BmJl = True Then
        txtYj1.Locked = False
    End If
ElseIf mod1.BmJl = True And lblLc.Caption = 2 And mod1.DName = lblLcRen.Caption Then
    frmFX.Visible = True
    dt3.Enabled = True
    dt4.Enabled = True
    optLa.Enabled = True
    optLb.Enabled = True
    optLc.Enabled = True
    cmdSave.Enabled = True
    txtHtze.Locked = False
    txtFbnr.Locked = False
    txtWBNR.Locked = False
    cmdYadd.Visible = True
    cmdYdel.Visible = True
    txtBz.Locked = False
    txtYj1.Locked = False
    If comQy.Text <> "�Ϻ�" And Val(lblMHid.Caption) < 19345 Then
        txtW3.Locked = False
        txtW4.Locked = False
    End If
    txtW5.Locked = False
    txtW6.Locked = False

    comYjRen.Locked = False
    tt = "SELECT dbo.khRen.khMan, dbo.khRen.rId FROM dbo.khRen INNER JOIN dbo.khzl ON dbo.khRen.khDh = dbo.khzl.khDh where dbo.khRen.khDh='" & txtKhdm.Text & "' and khren.lc=100"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2) + 1
    If La > 0 Then
        For oo = 0 To La
            FMXC.comYjRen.AddItem Ra(0, oo)
            Rid(oo) = Ra(1, oo)
        Next
    End If
    '''''frmFk.Visible = True
ElseIf (mod1.DName = "����" Or mod1.DName = "������" Or mod1.DName = "������1") And optW.Value = False And optZ.Value = False Then
    frmPL.Visible = True
    frmFX.Visible = True
    dt3.Enabled = True
    dt4.Enabled = True
    optLa.Enabled = True
    optLb.Enabled = True
    optLc.Enabled = True
    txtBz.Locked = False
    txtTcBe.Locked = False
    txtHtze.Locked = False
    cmdSave.Enabled = True
    If comQy.Text <> "�Ϻ�" And Val(lblMHid.Caption) < 19345 Then
        txtW3.Locked = False
        txtW4.Locked = False
    End If
    txtFbnr.Locked = False
    txtWBNR.Locked = False
'''''    If lblyjFF.Caption = "False" Then
'''''        cmdYadd.Visible = True
'''''        cmdYdel.Visible = True
'''''    End If
    txtW5.Locked = False
    txtW6.Locked = False
    ''''frmFk.Visible = True
    'JILI = 0
ElseIf mod1.DName = "������" Then
    frmPL.Visible = True
    frmFX.Visible = True
    dt3.Enabled = True
    dt4.Enabled = True
    optLa.Enabled = True
    optLb.Enabled = True
    optLc.Enabled = True
    txtBz.Locked = False
    txtTcBe.Locked = False
    txtHtze.Locked = False
    cmdSave.Enabled = True
    txtYj1.Locked = False
    If comQy.Text <> "�Ϻ�" And Val(lblMHid.Caption) < 19345 Then
        txtW3.Locked = False
        txtW4.Locked = False
    End If
    txtFbnr.Locked = False
    txtWBNR.Locked = False
    If lblyjFF.Caption = "False" Then
        cmdYadd.Visible = True
        cmdYdel.Visible = True
    End If
    txtW3.Locked = False
    txtW4.Locked = False
    txtW5.Locked = False
    txtW6.Locked = False
    comKQY.Locked = False
    frmFk.Visible = True
'''''''ElseIf optZ.Value = True And mod1.BmJl = True And cmdMQm(1).Caption = mod1.DName Then
'''''''    cmdYadd.Visible = True
'''''''    cmdYdel.Visible = True
ElseIf mod1.DName = "�Ǽ���" Or mod1.DName = "�Ǽ���" Then
    comKQY.Locked = False
End If
cmdDel.Enabled = True
End Sub

Private Sub cmdNew_Click()
Dim W1 As Single
Dim W2 As Single
Dim W3 As Single
Dim W5 As Single
Dim W6 As Single
Dim FPLX As String







timZm = 17 '������ǩ��ͬ
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "����"
    mod1.cmd.Parameters("@bh") = Val(lblMHid.Caption)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mlt3") = "" 'ҵ����ע
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = 0
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
    mod1.cmd.Parameters("@mm19") = 0
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If

    
Set mod1.cmd = Nothing
End Sub

Private Sub cmdPje_Click()
Dim tt As String
On Error Resume Next
Pje.Show
Set Pje.adoPje = CreateObject("adodb.recordset")
tt = "select trq,ywy,zn,bz,tf from pizu where (bh='" & lblMHid.Caption & "' and yid=80) order by pid desc"
Pje.adoPje.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText

Dim Ra: Dim La
Dim ii As Integer: Dim oo As Integer
Ra = Pje.adoPje.GetRows
Pje.adoPje.Close
Set Pje.adoPje = Nothing
La = UBound(Ra, 2): Pje.dtgPje.Rows = La + 20
Pje.dtgPje.Clear
For oo = 1 To La + 1
    Pje.dtgPje.Row = oo
    For ii = 1 To 6
        Pje.dtgPje.Col = ii
        Pje.dtgPje.Text = Ra(ii - 1, oo - 1)
        If ii = 5 Then
            If Pje.dtgPje.Text = "True" Then
                Pje.dtgPje.Text = "ͬ��"
            ElseIf Pje.dtgPje.Text = "False" Then
                Pje.dtgPje.Text = "����"
            End If

        End If
    Next
Next
For oo = 1 To La + 1
    Pje.dtgPje.Row = oo
    Pje.dtgPje.Col = 5
            If Pje.dtgPje.Text = "����" Then
                For ii = 1 To 5
                    Pje.dtgPje.Col = ii
                    Pje.dtgPje.CellForeColor = &HFF&
                Next
            End If
Next
Pje.dtgPje.Row = 0
Pje.dtgPje.Col = 1: Pje.dtgPje.Text = "����": Pje.dtgPje.Col = 2: Pje.dtgPje.Text = "����": Pje.dtgPje.Col = 3: Pje.dtgPje.Text = "ְ��"
Pje.dtgPje.Col = 4: Pje.dtgPje.Text = "������": Pje.dtgPje.Col = 5: Pje.dtgPje.Text = "ͨ����"
Pje.dtgA.Clear
Pje.dtgA.Rows = Pje.dtgPje.Rows
Pje.dtgA.Cols = Pje.dtgPje.Cols
For oo = 0 To Pje.dtgPje.Rows
    Pje.dtgPje.Row = oo
    Pje.dtgA.Row = oo
    For ii = 0 To Pje.dtgPje.Cols
        Pje.dtgPje.Col = ii
        Pje.dtgA.Col = ii
        Pje.dtgA.Text = Pje.dtgPje.Text
    Next
Next
End Sub

Private Sub cmdQing_Click()
txtYed.Text = ""
End Sub

Private Sub cmdMQm_Click(Index As Integer)

Dim ZED As Double
Dim oo As Integer
Dim Zje As Double
Dim tt As String
Dim Tywy As String '������ת����һ�˵�����
Dim Tuid As String
Dim Oywy As String 'ԭ����ת�˵�����
Dim Ouid As String 'ԭ����ת�˵Ĺ���
Dim Bid1 As Long
Dim Bid6 As Long
Dim Bid7 As Long
On Error Resume Next
optC.Visible = False
If Index = 5 And (mod1.DName = "�Ǽ���" Or mod1.DName = "�Ǽ���" Or mod1.DName = "�Ǽ���") Then
    frmQm.Visible = True
    OptT1.Value = False
    optT2.Value = False
    optC.Visible = True
    Exit Sub
End If
'�ɰ治��������
If mod1.ZT = "HMData" And (lblLc.Caption = 1 Or lblLc.Caption = 0) And lblLcRen.Caption = mod1.DName Then
    dtgFL.Col = 4
    dtgFL.Row = 1: Bid1 = Val(Mid(dtgFL.Text, 4, Len(dtgFL.Text) - 3))
    dtgFL.Row = 6: Bid6 = Val(Mid(dtgFL.Text, 4, Len(dtgFL.Text) - 3))
    dtgFL.Row = 7: Bid7 = Val(Mid(dtgFL.Text, 4, Len(dtgFL.Text) - 3))
    oo = MsgBox("��Ϊ�ɰ��ͬ,��Ҫ������,������Ϊ�°�,�Ƿ�ͨ��������������?", vbInformation + vbYesNo, "����!")
    If oo = vbYes Then
        timZm = 20 '�汾����
            Set mod1.cmd = CreateObject("adodb.command")
            mod1.cmd.ActiveConnection = mod1.cc
            mod1.cmd.CommandText = "MLAdd"
            mod1.cmd.CommandType = adCmdStoredProc
            mod1.cmd.Parameters("@zid") = 0
            mod1.cmd.Parameters("@errch") = ""
            mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
            mod1.cmd.Parameters("@NBLX") = "�汾����"
            mod1.cmd.Parameters("@bh") = Val(lblMHid.Caption)
            mod1.cmd.Parameters("@ywy") = mod1.DName
            mod1.cmd.Parameters("@uid") = mod1.DHid
            mod1.cmd.Parameters("@mt1") = lblHtxz.Caption
            mod1.cmd.Parameters("@mlt1") = ""
            mod1.cmd.Parameters("@mm1") = Bid1
            mod1.cmd.Parameters("@mm6") = Bid6
            mod1.cmd.Parameters("@mm7") = Bid7
            mod1.cmd.Parameters("@mb1") = 0
            Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
            mod1.cmd.Execute
            mod1.Zid = mod1.cmd.Parameters("@zid").Value
            If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
                MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
                cmdDing.Enabled = False
                Exit Sub
            Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
                Me.Enabled = False
                frmWaitA.Visible = True
                frmWaitA.Timer2.Enabled = False
                
                frmWaitA.ZOrder 0
                frmWaitA.Timer2.Enabled = True
                timWait.Enabled = True
            End If
        
            
        Set mod1.cmd = Nothing
    Else
        Exit Sub
    End If
    Exit Sub
End If





If Me.Visible = False Then Exit Sub
If cmdMQm(Index).Caption <> "" Then
    Exit Sub
End If

If Val(txtHtbh.ToolTipText) = 0 And Val(lblLc.Caption) = 1 Then
    MsgBox "�뵼����Ӱ��ͬ!"
    Call txtHtbh_DblClick
    Exit Sub
End If


If comFP.Text = "" Then
    MsgBox ("��ѡ��Ʊ��ʽ!")
    cmdSave.Enabled = True
    Exit Sub
End If

If lblHtxz.Caption = "ά��" And (txtF.Text = "" Or txtL.Text = "") Then
    MsgBox ("�����ά������ʼ�ںͽ�����!")
    frmWbNew.tabHt.Tab = 1
    cmdSave.Enabled = True
    Exit Sub
End If

'����Ӧ�տ������ܽ����һ��
MMdtgFk.Col = 3
Zje = 0
For oo = 0 To MMdtgFk.Rows - 1
    MMdtgFk.Row = oo
    Zje = Zje + Val(MMdtgFk.Text)
Next
'''''''''''''If Val(Zje) <> Val(txtHtze.Text) Then
'''''''''''''    If Val(lblLc.Caption) > 1 Then
'''''''''''''        txtQM.Text = "�տ���ϸ�����տһ�£���ȷ��"
'''''''''''''        frmQm.Visible = True
'''''''''''''        OptT1.Enabled = False
'''''''''''''        optT2.Enabled = True
'''''''''''''        optT2.Value = True
'''''''''''''    Else
'''''''''''''        MsgBox "�տ���ϸ�����տһ�£���ȷ��"
'''''''''''''    End If
'''''''''''''
'''''''''''''    Exit Sub
'''''''''''''End If
'''''''''''''ZED = Zje
'''''''''''''
'''''''''''''MMdtgFk.Col = 2
'''''''''''''Zje = 0
'''''''''''''For oo = 1 To 20
'''''''''''''    MMdtgFk.Row = oo
'''''''''''''    Zje = Zje + Val(MMdtgFk.Text)
'''''''''''''Next
'''''''''''''If Round(Zje, 0) <> 100 And ZED <> Val(txtHtze.Text) Then
'''''''''''''    MsgBox ("�����븶�ʽ!")
'''''''''''''    cmdSave.Enabled = True
'''''''''''''    Exit Sub
'''''''''''''End If



If cmdSave.Enabled = True Then
    MsgBox "���Ƚ����ӱ���,��ǩ�����Ĵ���!"
    Exit Sub
End If

If Index + 1 <> lblLc.Caption And lblLc.Caption <> 0 Then '�����ڲ���ɵ�λ�����ҵ�
    Exit Sub
End If

If mod1.DName = "����" And lblLcRen.Caption = "��ά" Then
    lblLcRen.Caption = "����"
    lblLcUid.Caption = "HM154"
End If

If lblLcUid.Caption <> mod1.DHid Then
'''    If lblLc.Caption = 3 And mod1.DName = "�Ǽ���" Then
'''    Else
        If Not (Val(lblLc.Caption) = 1 And txtXYwy.Text = mod1.DName) Then
            MsgBox "�˴�Ӧ��" & lblLcRen.Caption & "ǩ��! ������Ҫ�ٵ�"
            Exit Sub
        End If
'''    End If
End If


If txtHtbh.Text = "HMNEW" Then
    MsgBox ("�������ɺ�ͬ���!")
    Exit Sub
End If

If optY1.Value = False And optY2.Value = False Then
    MsgBox ("��ȷ���Ƿ������Ŀ����!")
    Exit Sub
End If



Dim W5 As Single
Dim W6 As Single
If Val(txtH5.Text) > 0 Then
    W5 = Val(txtH5.Text)
Else
    W5 = Val(txtW5.Text)
End If
If Val(txtH6.Text) > 0 Then
    W6 = Val(txtH6.Text)
Else
    W6 = Val(txtW6.Text)
End If

If Val(txtClcb1.Text) <> Val(W5 + W6) Then
MsgBox "���ϳɱ������밴�ύ��ť�����¼���ɱ�"
cmdSave.Enabled = True
Exit Sub
End If

If lblMQM(Index).Caption = "ҵ��Ա" And Val(txtHtze.Text) >= 15000 Then
    If (Val(txtW5.Text) > 0 And Val(txtH5.Text) = 0) Or (Val(txtW6.Text) > 0 And Val(txtH6.Text) = 0) Then
        MsgBox ("����15000�ĺ�ͬ������ϳɱ��������Ⱦ���ѯ�۵�����ʽ�˼ۣ� ")
        Exit Sub
    End If
End If

frmQm.Visible = True
If Index = 0 Then 'ҵ��Աֻ��ǩ�֣����ܲ��ء�
    optT2.Enabled = False
    OptT1.Enabled = True
Else
    optT2.Enabled = True
End If
OptT1.Value = True

If lblMQM(Index).Caption = "��ִͬ��" Then
    If (Val(txtW5.Text) > 0 And Val(txtH5.Text) = 0) Or (Val(txtW6.Text) > 0 And Val(txtH6.Text) = 0) Then
        MsgBox ("���ϳɱ����ں�ִͬ��ǰ�����뾭��ѯ�۵�����ʽ�˼ۣ� �뽫�˵�����")
        OptT1.Enabled = False
        optT2.Enabled = True
        optT2.Value = True
        Exit Sub
    End If
End If

'If lblMQM(Index).Caption = "���ȷ��" And chkQKF.Value = 0 Then
'
'        MsgBox ("δ��ȫ����ܵ���ɣ�")
'        frmQM.Visible = False
'        Exit Sub
'
'End If
Exit Sub







End Sub

Private Sub cmdQkfAdd_Click()
Dim tt As String
Dim hg As Single
On Error Resume Next
If Val(txtQkfJe.Text) = 0 Then
Exit Sub
End If

timZm = 16 'ȫ��༭
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "ȫ��༭"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text) '��ͬ���
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mlt1") = txtQkFBz.Text
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtQkfJe.Text)
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 1 '���ȫ��
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = dtpQkF.Value
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


Set mod1.cmd = Nothing
End Sub

Private Sub cmdQkfDel_Click()
Dim tt As String
Dim ii As Integer
Dim Yid As Long

On Error Resume Next

dtgyjF.Col = 4
Yid = 0
Yid = Val(dtgyjF.Text)


If Yid = 0 Then
Exit Sub
End If



ii = MsgBox("�Ƿ�ɾ���˼�¼?", vbQuestion + vbYesNo, "ѯ��")
If ii = vbNo Then
    Exit Sub
End If

timZm = 16 'ȫ��༭
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "ȫ��༭"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text) '��ͬ���
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mm1") = Yid
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 0 'ȫ��ɾ��
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


Set mod1.cmd = Nothing
End Sub

Private Sub cmdSave_Click()
Dim W1 As Single
Dim W2 As Single
Dim W3 As Single
Dim W4 As Single
Dim W5 As Single
Dim W6 As Single
Dim W7 As Single '��פ
Dim FPLX As String


'����˼۳ɱ�����Ԥ�Ƴɱ�,��ȡ�˼۳ɱ�,�����ȡԤ�Ƴɱ��������.

'''''''''If Val(txtYj1.Text) > 0 And (comYjRen.Text = "" Or comYjRen.ToolTipText = "") And Val(lblLc.Caption) > 1 Then'�������޸�
'''''''''    MsgBox "û��ȷ�����˱�����"
'''''''''    Exit Sub
'''''''''End If
If Val(FmxcFK.txtBL1.Text) > 100 Or Val(FmxcFK.txtBL1.Text) < 0 Or Val(FmxcFK.txtBL2.Text) > 100 Or Val(FmxcFK.txtBL2.Text) < 0 Or Val(FmxcFK.txtBL3.Text) > 100 Or Val(FmxcFK.txtBL3.Text) < 0 Then
    MsgBox "û����ȷ���ÿ�����ɱ���!"
    Exit Sub
    
End If
dtgFL.Col = 2: dtgFL.Row = 1: W1 = Val(dtgFL.Text)
dtgFL.Col = 2: dtgFL.Row = 2: W2 = Val(dtgFL.Text)
dtgFL.Col = 2: dtgFL.Row = 3: W3 = Val(dtgFL.Text)
dtgFL.Col = 2: dtgFL.Row = 4: W4 = Val(dtgFL.Text)
dtgFL.Col = 2: dtgFL.Row = 6: W5 = Val(dtgFL.Text)
dtgFL.Col = 2: dtgFL.Row = 7: W6 = Val(dtgFL.Text)
dtgFL.Col = 2: dtgFL.Row = 5: W7 = Val(dtgFL.Text)


    W2 = Val(txtH2.Text)



''''''''If Val(txtH3.Text) > 0 Then
''''''''    W3 = Val(txtH3.Text)
''''''''Else
'''''''    W3 = Val(txtW3.Text)
''''''''End If
'''''''If Val(txtH5.Text) > 0 Then
'''''''    W5 = Val(txtH5.Text)
'''''''Else
'''''''    W5 = Val(txtW5.Text)
'''''''End If
'''''''If Val(txtH6.Text) > 0 Then
'''''''    W6 = Val(txtH6.Text)
'''''''Else
'''''''    W6 = Val(txtW6.Text)
'''''''End If

txtRgf1.Text = W1 + W2
txtFbje1.Text = W3 + W4 + W7
txtClcb1.Text = W5 + W6

If FMXC.FO = 0 Then FMXC.FO = 1
If lblHtxz.Caption = "ά��" Or lblHtxz.Caption = "ˮ����" Then
'����ɱ�����
    txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + Val(txtYf1.Text)
    txtCbze1.Text = Round(Val(txtCbze1.Text) / FMXC.FO, 2)
    txtJlr1.Text = Val(txtHtze.Text) - Val(txtCbze1.Text)
    txtLr1.Text = Val(txtJlr1.Text) - Val(txtYj1.Text)
'''''''''''''''    txtQt1.Text = Val(txtLr1.Text) * 0.1
'''''''''''''''
'''''''''''''''    txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + Val(txtYf1.Text) + Val(txtQt1.Text)
'''''''''''''''    txtJlr1.Text = Val(txtHtze.Text) - Val(txtCbze1.Text)
'''''''''''''''    txtLr1.Text = Val(txtJlr1.Text) - Val(txtYj1.Text)
Else
    txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + Val(txtYf1.Text)
    txtCbze1.Text = Round(Val(txtCbze1.Text) / FMXC.FO, 2)
    txtJlr1.Text = Val(txtHtze.Text) - Val(txtCbze1.Text)
    txtLr1.Text = Val(txtJlr1.Text) - Val(txtYj1.Text)
End If

''''If optLa.Value = True Then
''''    FPLX = "��ֵ��Ʊ"
''''ElseIf optLb.Value = True Then
''''    FPLX = "��ҵ��Ʊ"
''''ElseIf optLc.Value = True Then
''''    FPLX = "����Ʊ"
''''End If
FPLX = comFP.Text
If txtTcRQ.Text = "" Then
    txtTcRQ.Text = "2000-1-1"
End If


Call DJ '�����ٴ���

timZm = 2 '�����ͬ
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "����"
    mod1.cmd.Parameters("@bh") = Val(lblMHid.Caption)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = FPLX '��Ʊ����
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mt25") = comKQY.Text '��������
    mod1.cmd.Parameters("@mt26") = FmxcFK.comQy3.Text '��������
    mod1.cmd.Parameters("@mt27") = FmxcFK.txtRen2.Text '��������
    mod1.cmd.Parameters("@mt28") = FmxcFK.txtRen3.Text  '��������
    mod1.cmd.Parameters("@mt29") = FmxcFK.txtRen2.ToolTipText  '��������
    mod1.cmd.Parameters("@mt30") = FmxcFK.txtRen3.ToolTipText  '��������
    mod1.cmd.Parameters("@mt31") = FmxcFK.txtBL1.Text '��������
    mod1.cmd.Parameters("@mt32") = FmxcFK.txtBL2.Text   '��������
    mod1.cmd.Parameters("@mt33") = FmxcFK.txtBL3.Text  '��������
    mod1.cmd.Parameters("@mlt1") = txtBz.Text '��ע
    'mod1.cmd.Parameters("@mlt2") = txtWBNR.Text '�������
    mod1.cmd.Parameters("@mlt3") = "" 'ҵ����ע
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtYf1.Text) '�˷�
    mod1.cmd.Parameters("@mm2") = Val(txtTcBe.Text) '��ɱ���
    mod1.cmd.Parameters("@mm3") = Val(lblLc.Caption) '�������Ϊ0,�����ҵ��Ա������
    mod1.cmd.Parameters("@mm4") = FMXC.FO '���ʽ����
    mod1.cmd.Parameters("@mm5") = Val(txtW3.Text)
    mod1.cmd.Parameters("@mm6") = Val(txtW4.Text)
    mod1.cmd.Parameters("@mm7") = Val(txtW5.Text)
    mod1.cmd.Parameters("@mm8") = Val(txtW6.Text)
    mod1.cmd.Parameters("@mm9") = Val(txtCbze1.Text)
    mod1.cmd.Parameters("@mm10") = Val(txtClcb1.Text)
    mod1.cmd.Parameters("@mm11") = Val(txtRgf1.Text)
    mod1.cmd.Parameters("@mm12") = Val(txtCLF1.Text)
    mod1.cmd.Parameters("@mm13") = Val(txtFbje1.Text)
    mod1.cmd.Parameters("@mm14") = 0
    mod1.cmd.Parameters("@mm15") = Val(txtQt1.Text)
    mod1.cmd.Parameters("@mm16") = Val(txtJlr1.Text)
    mod1.cmd.Parameters("@mm17") = Val(txtLr1.Text)
    mod1.cmd.Parameters("@mm18") = Val(txtHtze.Text)
    mod1.cmd.Parameters("@mm22") = Val(comYjRen.ToolTipText)
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mm21").Value = Val(txtYj1.Text)
    dtgFL.Col = 3: dtgFL.Row = 1
    mod1.cmd.Parameters("@mm23") = Val(dtgFL.Text)  '�ٴ��� ά��
    dtgFL.Col = 3: dtgFL.Row = 2
    mod1.cmd.Parameters("@mm24") = Val(dtgFL.Text)  '�ٴ��� ����
    dtgFL.Col = 3: dtgFL.Row = 3
    mod1.cmd.Parameters("@mm25") = Val(dtgFL.Text)  '�ٴ��� ���̷ְ�
    dtgFL.Col = 3: dtgFL.Row = 4
    mod1.cmd.Parameters("@mm26") = Val(dtgFL.Text)  '�ٴ��� ˮ����
    dtgFL.Col = 3: dtgFL.Row = 5
    mod1.cmd.Parameters("@mm27") = Val(dtgFL.Text)  '�ٴ��� ��פ
    dtgFL.Col = 3: dtgFL.Row = 6
    mod1.cmd.Parameters("@mm28") = Val(dtgFL.Text)  '�ٴ��� ���
    dtgFL.Col = 3: dtgFL.Row = 7
    mod1.cmd.Parameters("@mm29") = Val(dtgFL.Text)  '�ٴ��� ��Ʒ
    dtgFL.Col = 2: dtgFL.Row = 5
    mod1.cmd.Parameters("@mm30") = Val(dtgFL.Text)  '��פ��׼��
    mod1.cmd.Parameters("@mb1") = chkYJF.Value 'ҵ����
    mod1.cmd.Parameters("@mb2") = chkJTF.Value '��ɷ�
    mod1.cmd.Parameters("@mb3") = chkQKF.Value 'ȫ���
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = FMXC.dt3.Value 'ά����ʼ��
    mod1.cmd.Parameters("@md2") = FMXC.dt4.Value
    mod1.cmd.Parameters("@md3") = FMXC.txtHtrq.Text '��ͬ����
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        If timZm = 2 Then '����
            cmdSave.Enabled = False
        End If
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
        frmFk.Visible = False
        frmFX.Visible = False
        
    End If

    
Set mod1.cmd = Nothing


End Sub




Private Sub cmdW1_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next

'''''frmWBXNew.Show
'''''Exit Sub


'''''Call frmWBXNew.Qing
'''''Call frmWBXNew.Bound(Val(cmdW1.ToolTipText))
'''''frmWBXNew.Show
'''''Exit Sub
Me.OldF = True

If Val(cmdW1.ToolTipText) > 0 Then
    mod1.BTZ = 36
    'If Val(cmdW1.ToolTipText) > 8052 Then
    If Me.OldF = True Then

        Call frmWBXNew.Qing
        Call frmWBXNew.Bound(cmdW1.ToolTipText)
        frmWBXNew.Show
        frmWBXNew.ZOrder 0

        Exit Sub
    'End If
    Else
        Call modBJD.BJDWBQing
        Call modBJD.BJDBound(cmdW1.ToolTipText, "ά��")
        Call modBJD.wbxjLocked
        frmWBXJ.Show
        frmWBXJ.lblLcUid.Caption = FMXC.txtXYwy.ToolTipText
        frmWBXJ.lblLcRen.Caption = FMXC.txtXYwy.Text
        Exit Sub
    End If
End If

If Val(cmdW1.ToolTipText) = 0 And (txtYwy.ToolTipText = mod1.DHid Or mod1.DName = "" Or mod1.DName = "�ܴ���") And txtHtbh.Text = "HMNEW" Then
    If mod1.DName <> txtYwy.Text Or lblLc.Caption > 1 Then
    Exit Sub
    End If
    ii = MsgBox("�Ƿ��½�ά��ѯ�۵�?", vbInformation + vbYesNo, "Hello!")
'''    MsgBox ("���ڲ����У�����϶����ã����½⣡")
'''    Exit Sub
    If ii = vbNo Then Exit Sub
   
    
    frmWBXJ.Visible = False
    Call modBJD.BJDWBQing
    Call modBJD.wbxjUnLocked
    
    
timZm = 3 '�½�ѯ�۵�
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "�½�ѯ�۵�"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "ά��"
    mod1.cmd.Parameters("@mt2") = txtXmmc.Text
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mm1") = 88 'NLBֵ
    mod1.cmd.Parameters("@mm2") = txtXmmc.ToolTipText '��Ŀ���
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    LLXX = True
    mod1.cmd.Parameters("@mb1") = 1 'LXֵ
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"

        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
    
    mod1.BTZ = 36
End If

Exit Sub



If mod1.DName <> txtYwy.Text Or lblLc.Caption > 1 Then
Exit Sub
End If
ii = MsgBox("�Ƿ��½�ά��ѯ�۵�?", vbInformation + vbYesNo, "Hello!")
If ii = vbNo Then Exit Sub

Me.Enabled = False
mod1.BTZ = 36
'���½�ά��ѯ��
frmWBXJ.Visible = False
Call modBJD.BJDWBQing
Call modBJD.wbxjUnLocked
frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.workKK
mod1.cmd.CommandText = "xunJiaAddHT"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@ywy") = mod1.DName
mod1.cmd.Parameters("@uid") = mod1.DHid
mod1.cmd.Parameters("@Lx") = 1
mod1.cmd.Parameters("@zl") = "ά��"
mod1.cmd.Parameters("@Lcou") = 4 '��������
mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
mod1.cmd.Parameters("@lcRen") = mod1.DName
mod1.cmd.Parameters("@lcUid") = mod1.DHid
mod1.cmd.Parameters("@nLb") = 44
mod1.cmd.Parameters("@xmmc") = txtXmmc.Text
mod1.cmd.Parameters("@xid") = txtXmmc.ToolTipText
mod1.cmd.Parameters("@errch") = ""
mod1.cmd.Parameters("@htbh") = lblMHid.Caption
mod1.cmd.Execute
frmWBXJ.lblBid.Caption = mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblBh.Caption = "XJD" & mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblOid.Caption = mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblLcou.Caption = 4 '��������
frmWBXJ.lblLc.Caption = 0
frmWBXJ.lblLcRen.Caption = mod1.DName
frmWBXJ.lblLcUid.Caption = mod1.DHid
frmWBXJ.lblNlb.Caption = 44
frmWBXJ.lblYwy.Caption = mod1.DName
frmWBXJ.lblUid.Caption = mod1.DHid
frmWBXJ.lblBM.Caption = mod1.Bm
frmWBXJ.lblQy.Caption = mod1.Qy
frmWBXJ.lblZl.Caption = "ά��"
Set cmd = Nothing
If frmWBXJ.lblBh.Caption = "" Then
    ii = MsgBox("ϵͳ������������,�����̹ر�!�ٴδ򿪺�����Ϣ,������˴���.", vbOKOnly + vbExclamation, "A������")
    End
End If


'tt = "select jzpb,pbid from bjxt_jzpb"
'frmWBXJ.adoPb.Close
'frmWBXJ.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Set frmWBXJ.comPb.RowSource = frmWBXJ.adoPb
'frmWBXJ.comPb.ListField = "jzpb"
'frmWBXJ.comPb.BoundColumn = "pbid"
            frmWBXJ.frmDx.Visible = False
            frmWBXJ.frmNb.Visible = True
            frmWBXJ.frmTime.Visible = True

            frmWBXJ.cmdD.Visible = True
            frmWBXJ.cmdJi.Visible = True
            frmWBXJ.tabGc.TabVisible(2) = False
            frmWBXJ.tabGc.TabVisible(0) = True
            frmWBXJ.tabGc.TabVisible(1) = True
            frmWBXJ.tabGc.Tab = 0

    '�������̰�ť
    Call modBJD.XJWBLcBut(44)
    
        frmWBXJ.cmdD.Visible = True

        frmWBXJ.cmdJi.Visible = True
    
frmWait.Visible = False
frmWBXJ.Visible = True
frmWBXJ.cmdMod.Enabled = False
''ˢ��ά�������б�
'tt = "select * from xunJIaWbView where wbx='�걣' and bid=" & Val(frmWBXJ.lblBid.Caption)
'    frmWBXJ.adoWb.Close
'    frmWBXJ.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    Set frmWBXJ.dtgWb.DataSource = frmWBXJ.adoWb
'tt = "select * from xunJIaWbView where wbx='����' and bid=" & Val(frmWBXJ.lblBid.Caption)
'    frmWBXJ.adoLj.Close
'    frmWBXJ.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    Set frmWBXJ.dtgLj.DataSource = frmWBXJ.adoLj
'    frmWBXJ.cmdSave.Enabled = True
'frmGxBiao.Enabled = False

'������Ϣ��
frmWBXJ.frmNew.Visible = True
tt = "select jzpb as ����Ʒ��,jzxh as �����ͺ�,sl as ����,jxId from wbjb where bid=" & Val(frmWBXJ.lblBid.Caption)
Set mod1.mA = CreateObject("adodb.recordset")
mod1.mA.Close
mod1.mA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.dtgA.DataSource = mod1.mA


'���º�ͬ
tt = "update htping set bid1=" & Val(frmWBXJ.lblBid.Caption) & "where hid=" & Val(lblMHid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
cmdW1.ToolTipText = frmWBXJ.lblBid.Caption


frmWBXJ.cmdBjd.Visible = False
frmWBXJ.txtHg.Locked = True
frmWBXJ.txtYhg.Locked = True
frmWBXJ.txtClf.Locked = True
frmWBXJ.cmdCG.Enabled = False
'frmWBXJ.cmdCong.Visible = False
frmWBXJ.cmdTK.Visible = True
frmWBXJ.Visible = True
frmWBXJ.comXmmc.Text = txtXmmc.Text
frmWBXJ.comXmmc.ToolTipText = txtXmmc.ToolTipText
frmWBXJ.cmdSave.Enabled = True

End Sub


Private Sub cmdW2_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next

If Val(cmdW2.ToolTipText) > 0 Then
    mod1.BTZ = 36
    If Val(frmPrf.Caption) > 0 Then
        Call frmWBXNew.Qing
        Call frmWBXNew.Bound(cmdW2.ToolTipText)
        frmWBXNew.frmM1.Visible = False
        frmWBXNew.Show
        frmWBXNew.ZOrder 0
        Exit Sub
    End If

    Call modBJD.BJDWBQing
    Call modBJD.BJDBound(cmdW2.ToolTipText, "����")
    Call modBJD.wbxjLocked
    frmWBXJ.Show
    frmWBXJ.lblLcUid.Caption = FMXC.txtXYwy.ToolTipText
    frmWBXJ.lblLcRen.Caption = FMXC.txtXYwy.Text
    Exit Sub
End If

If (Val(cmdW2.ToolTipText) = 0 And (txtYwy.ToolTipText = mod1.DHid Or mod1.DName = "" Or mod1.DName = "�ܴ���") And txtHtbh.Text = "HMNEW") Or mod1.DName = "������" Then
    If (mod1.DName <> txtYwy.Text Or lblLc.Caption > 1) And mod1.DName <> "������" Then
    Exit Sub
    End If
    ii = MsgBox("�Ƿ��½�����ѯ�۵�?", vbInformation + vbYesNo, "Hello!")
'''''    MsgBox ("���ڲ����У�����϶����ã����½⣡")
'''''    Exit Sub
    If ii = vbNo Then Exit Sub
   
    
    frmWBXJ.Visible = False
    Call modBJD.BJDWBQing
    Call modBJD.wbxjUnLocked
    
    
timZm = 3 '�½�ѯ�۵�
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "�½�ѯ�۵�"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "����"
    mod1.cmd.Parameters("@mt2") = txtXmmc.Text
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mm1") = 88 'NLBֵ
    mod1.cmd.Parameters("@mm2") = txtXmmc.ToolTipText '��Ŀ���
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    LLXX = True
    mod1.cmd.Parameters("@mb1") = 1 'LXֵ
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"

        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
    
    mod1.BTZ = 36
End If

Exit Sub

If Val(cmdW2.ToolTipText) > 0 Then
mod1.BTZ = 36
Call modBJD.BJDWBQing
Call modBJD.BJDBound(cmdW2.ToolTipText, "����")
frmWBXJ.Show
frmWBXJ.cmdSave.Enabled = True
frmWBXJ.frmTime.Visible = False
frmWBXJ.frmNb.Visible = False
frmWBXJ.cmdD.Visible = False
frmWBXJ.cmdTK.Visible = False
frmWBXJ.cmdCG.Visible = False
Exit Sub
End If

If mod1.DName <> txtYwy.Text Or lblLc.Caption > 1 Then
Exit Sub
End If
ii = MsgBox("�Ƿ��½�����ѯ�۵�?", vbInformation + vbYesNo, "Hello!")
If ii = vbNo Then Exit Sub

Me.Enabled = False
mod1.BTZ = 36
'���½�ά��ѯ��
frmWBXJ.Visible = False
Call modBJD.BJDWBQing
Call modBJD.wbxjUnLocked
frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.workKK
mod1.cmd.CommandText = "xunJiaAddHT"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@ywy") = mod1.DName
mod1.cmd.Parameters("@uid") = mod1.DHid
mod1.cmd.Parameters("@Lx") = 1
mod1.cmd.Parameters("@zl") = "����"
mod1.cmd.Parameters("@Lcou") = 4 '��������
mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
mod1.cmd.Parameters("@lcRen") = mod1.DName
mod1.cmd.Parameters("@lcUid") = mod1.DHid
mod1.cmd.Parameters("@nLb") = 44
mod1.cmd.Parameters("@xmmc") = FMXC.txtXmmc.Text
mod1.cmd.Parameters("@xid") = FMXC.txtXmmc.ToolTipText
mod1.cmd.Parameters("@errch") = ""
mod1.cmd.Parameters("@htbh") = FMXC.lblMHid.Caption
mod1.cmd.Execute
frmWBXJ.lblBid.Caption = mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblBh.Caption = "XJD" & mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblOid.Caption = mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblLcou.Caption = 4 '��������
frmWBXJ.lblLc.Caption = 0
frmWBXJ.lblLcRen.Caption = mod1.DName
frmWBXJ.lblLcUid.Caption = mod1.DHid
frmWBXJ.lblNlb.Caption = 44
frmWBXJ.lblYwy.Caption = mod1.DName
frmWBXJ.lblUid.Caption = mod1.DHid
frmWBXJ.lblBM.Caption = mod1.Bm
frmWBXJ.lblQy.Caption = mod1.Qy
frmWBXJ.lblZl.Caption = "����"
Set cmd = Nothing
If frmWBXJ.lblBh.Caption = "" Then
    ii = MsgBox("ϵͳ������������,�����̹ر�!�ٴδ򿪺�����Ϣ,������˴���.", vbOKOnly + vbExclamation, "A������")
    End
End If
'������Ŀ������Ϣ
tt = "select xmmc,xid from xmzl where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"
frmWBXJ.adoXm.Close
frmWBXJ.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.comXmmc.RowSource = frmWBXJ.adoXm
frmWBXJ.comXmmc.ListField = "xmmc"
frmWBXJ.comXmmc.BoundColumn = "xid"

'tt = "select jzpb,pbid from bjxt_jzpb"
'frmWBXJ.adoPb.Close
'frmWBXJ.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Set frmWBXJ.comPb.RowSource = frmWBXJ.adoPb
'frmWBXJ.comPb.ListField = "jzpb"
'frmWBXJ.comPb.BoundColumn = "pbid"
            frmWBXJ.frmDx.Visible = False
            frmWBXJ.frmNb.Visible = True
            frmWBXJ.frmTime.Visible = True

            frmWBXJ.cmdD.Visible = True
            frmWBXJ.cmdJi.Visible = True
            frmWBXJ.tabGc.TabVisible(2) = True
            frmWBXJ.tabGc.TabVisible(0) = False
            frmWBXJ.tabGc.TabVisible(1) = False
            frmWBXJ.tabGc.Tab = 0

    '�������̰�ť
    Call modBJD.XJWBLcBut(44)
    
        frmWBXJ.cmdD.Visible = True

        frmWBXJ.cmdJi.Visible = True
    
frmWait.Visible = False
frmWBXJ.Visible = True
frmWBXJ.cmdMod.Enabled = False
''ˢ��ά�������б�
'tt = "select * from xunJIaWbView where wbx='�걣' and bid=" & Val(frmWBXJ.lblBid.Caption)
'    frmWBXJ.adoWb.Close
'    frmWBXJ.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    Set frmWBXJ.dtgWb.DataSource = frmWBXJ.adoWb
'tt = "select * from xunJIaWbView where wbx='����' and bid=" & Val(frmWBXJ.lblBid.Caption)
'    frmWBXJ.adoLj.Close
'    frmWBXJ.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    Set frmWBXJ.dtgLj.DataSource = frmWBXJ.adoLj
'    frmWBXJ.cmdSave.Enabled = True
'frmGxBiao.Enabled = False

'������Ϣ��
frmWBXJ.frmNew.Visible = True
tt = "select jzpb as ����Ʒ��,jzxh as �����ͺ�,sl as ����,jxId from wbjb where bid=" & Val(frmWBXJ.lblBid.Caption)
Set mod1.mA = CreateObject("adodb.recordset")
frmWBXJ.adoA.Close
frmWBXJ.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.dtgA.DataSource = frmWBXJ.adoA


'���º�ͬ
tt = "update htping set bid2=" & Val(frmWBXJ.lblBid.Caption) & "where hid=" & Val(lblMHid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
cmdW2.ToolTipText = frmWBXJ.lblBid.Caption


frmWBXJ.cmdBjd.Visible = False
frmWBXJ.txtHg.Locked = True
frmWBXJ.txtYhg.Locked = True
frmWBXJ.txtClf.Locked = True
frmWBXJ.cmdCG.Enabled = False
'frmWBXJ.cmdCong.Visible = False
frmWBXJ.cmdTK.Visible = True
frmWBXJ.Visible = True
frmWBXJ.comXmmc.Text = txtXmmc.Text
frmWBXJ.comXmmc.ToolTipText = txtXmmc.ToolTipText
frmWBXJ.cmdSave.Enabled = True
frmWBXJ.frmTime.Visible = False
frmWBXJ.frmNb.Visible = False
frmWBXJ.cmdD.Visible = False
frmWBXJ.cmdTK.Visible = False
frmWBXJ.cmdCG.Visible = False
frmWBXJ.txtDxnr.Locked = True
End Sub


Private Sub cmdW3_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next

If Val(cmdW3.ToolTipText) > 0 Then
    mod1.BTZ = 36
    If Val(frmPrf.Caption) > 0 Then
        Call frmWBXNew.Qing
        Call frmWBXNew.Bound(cmdW3.ToolTipText)
        frmWBXNew.Show
        frmWBXNew.ZOrder 0
        Exit Sub
    End If
    Call modBJD.BJDWBQing
    Call modBJD.BJDBound(cmdW3.ToolTipText, "���̷ְ�")
    Call modBJD.wbxjLocked
    frmWBXJ.Show
    frmWBXJ.lblLcUid.Caption = FMXC.txtXYwy.ToolTipText
    frmWBXJ.lblLcRen.Caption = FMXC.txtXYwy.Text
    Exit Sub
End If

If Val(cmdW3.ToolTipText) = 0 And (txtYwy.ToolTipText = mod1.DHid Or mod1.DName = "" Or mod1.DName = "�ܴ���") And (txtHtbh.Text = "HMNEW" Or (txtH1.Text <> "" Or txtH2.Text <> "") And Val(lblLc.Caption) = 1) Then
    If mod1.DName <> txtYwy.Text Or lblLc.Caption > 1 Then
        Exit Sub
    End If
'''''''    If comQy.Text <> "�Ϻ�" Then
'''''''        Exit Sub
'''''''    End If
    ii = MsgBox("�Ƿ��½����̷ְ�ѯ�۵�?", vbInformation + vbYesNo, "Hello!")
'''''    MsgBox ("���ڲ����У�����϶����ã����½⣡")
'''''    Exit Sub
    If ii = vbNo Then Exit Sub
   
    
    frmWBXJ.Visible = False
    Call modBJD.BJDWBQing
    Call modBJD.wbxjUnLocked
    
    
timZm = 3 '�½�ѯ�۵�
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "�½�ѯ�۵�"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "���̷ְ�"
    mod1.cmd.Parameters("@mt2") = txtXmmc.Text
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mm1") = 88 'NLBֵ
    mod1.cmd.Parameters("@mm2") = txtXmmc.ToolTipText '��Ŀ���
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    LLXX = True
    mod1.cmd.Parameters("@mb1") = 1 'LXֵ
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"

        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
    
    mod1.BTZ = 36
End If
End Sub


Private Sub cmdW4_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next

If Val(cmdW4.ToolTipText) > 0 Then
    mod1.BTZ = 36
    If Val(frmPrf.Caption) > 0 Then
        Call frmWBXNew.Qing
        Call frmWBXNew.Bound(cmdW4.ToolTipText)
        frmWBXNew.Show
        frmWBXNew.ZOrder 0
        Exit Sub
    End If
    Call modBJD.BJDWBQing
    Call modBJD.BJDBound(cmdW4.ToolTipText, "ˮ����")
    Call modBJD.wbxjLocked
    frmWBXJ.Show
    frmWBXJ.lblLcUid.Caption = FMXC.txtXYwy.ToolTipText
    frmWBXJ.lblLcRen.Caption = FMXC.txtXYwy.Text
    Exit Sub
End If

If Val(cmdW4.ToolTipText) = 0 And (txtYwy.ToolTipText = mod1.DHid Or mod1.DName = "" Or mod1.DName = "�ܴ���") And txtHtbh.Text = "HMNEW" Then
    If mod1.DName <> txtYwy.Text Or lblLc.Caption > 1 Then
    Exit Sub
    End If
'''    If comQy.Text <> "�Ϻ�" Then
'''        Exit Sub
'''    End If
    ii = MsgBox("�Ƿ��½�ˮ����ѯ�۵�?", vbInformation + vbYesNo, "Hello!")
'''''    MsgBox ("���ڲ����У�����϶����ã����½⣡")
'''''    Exit Sub
    If ii = vbNo Then Exit Sub


    frmWBXJ.Visible = False
    Call modBJD.BJDWBQing
    Call modBJD.wbxjUnLocked
    
    
timZm = 3 '�½�ѯ�۵�
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "�½�ѯ�۵�"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "ˮ����"
    mod1.cmd.Parameters("@mt2") = txtXmmc.Text
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mm1") = 88 'NLBֵ
    mod1.cmd.Parameters("@mm2") = txtXmmc.ToolTipText '��Ŀ���
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    LLXX = True
    mod1.cmd.Parameters("@mb1") = 1 'LXֵ
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"

        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
    
    mod1.BTZ = 36
End If
End Sub


Private Sub cmdW5_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
'If mod1.DName = mod1.DName Then
If Val(cmdW5.ToolTipText) = 0 And (txtYwy.ToolTipText = mod1.DHid Or mod1.DName = "" Or mod1.DName = "�ܴ���") And txtHtbh.Text = "HMNEW" Then
If mod1.DName <> txtYwy.Text Or lblLc.Caption > 1 Then
Exit Sub
End If
ii = MsgBox("�Ƿ��½����ѯ�۵�?", vbInformation + vbYesNo, "Hello!")
If ii = vbNo Then Exit Sub
    frmGXBj.Visible = False
    tt = "select jzpb,pbid from bjxt_jzpb"
    frmGXBj.adoPb.Close
    frmGXBj.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If IsNull(frmGXBj.adoPb.RecordCount) = True Then
        MsgBox ("��ȡ���ݳ���!")
        Exit Sub
    End If
    Set frmGXBj.comJzpb.RowSource = frmGXBj.adoPb
    frmGXBj.comJzpb.ListField = "jzpb"
    frmGXBj.comJzpb.BoundColumn = "pbid"
    
    
    frmGXBj.Visible = False
    Call modBJD.BJDGXQing
    Call modBJD.gxbjUnLocked

'    Set mod1.cmd = createobject("adodb.command")
'    mod1.cmd.ActiveConnection = mod1.CC
'    mod1.cmd.CommandText = "xunJiaAddHT"
'    mod1.cmd.CommandType = adCmdStoredProc
'    mod1.cmd.Parameters("@ywy") = mod1.DName
'    mod1.cmd.Parameters("@uid") = mod1.DHid
'    mod1.cmd.Parameters("@Lx") = 0
'    mod1.cmd.Parameters("@zl") = "����"
'    mod1.cmd.Parameters("@Lcou") = Right(frmGxBiao.cmdCreat.ToolTipText, 1) '��������
'    mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
'    mod1.cmd.Parameters("@lcRen") = mod1.DName
'    mod1.cmd.Parameters("@lcUid") = mod1.DHid
'    mod1.cmd.Parameters("@nLb") = 43
'    mod1.cmd.Parameters("@xmmc") = txtXMMC.Text
'    mod1.cmd.Parameters("@xid") = txtXMMC.ToolTipText
'    mod1.cmd.Parameters("@errch") = ""
'
'    mod1.cmd.Execute


    
timZm = 3 '�½�ѯ�۵�
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "�½�ѯ�۵�"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "���"
    mod1.cmd.Parameters("@mt2") = txtXmmc.Text
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mm1") = 43 'NLBֵ
    mod1.cmd.Parameters("@mm2") = txtXmmc.ToolTipText '��Ŀ���
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    LLXX = False
    mod1.cmd.Parameters("@mb1") = 0 'LXֵ
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        If timZm = 3 Then '����

        End If
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
    
    mod1.BTZ = 36
Else '�����ѯ�۵�
    If Val(cmdW5.ToolTipText) = 0 Then Exit Sub
    Call modBJD.BJDGXQing
    Call modBJD.BJDBound(Val(cmdW5.ToolTipText), "���")
    Call frmGXBj.SDJE(Val(txtD5.Text)) '��̯�ٴ���

    Call modBJD.gxbjLocked
    frmGXBj.optW.Value = True
    mod1.BTZ = 36
    frmWait.Visible = False
    frmGXBj.Visible = True
    frmGXBj.ZOrder 0
    frmGXBj.cmdMod.Enabled = True
    frmGXBj.cmdSave.Enabled = False
''''''    Pje.Visible = False
''''''    tt = "select * from pizu where bh='" & Val(cmdW5.ToolTipText) & "' and yid=43 order by trq desc"
''''''    Set Pje.adoPje = CreateObject("adodb.recordset")
''''''    Pje.adoPje.Close
''''''    Pje.adoPje.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''    If Pje.adoPje.RecordCount > 0 And Pje.adoPje.Fields("tf").Value = False Then
''''''         Set Pje.dtgPje.DataSource = Pje.adoPje
''''''        Pje.Visible = True
''''''        Pje.ZOrder 0
''''''        Pje.txtXQ.Text = ""
''''''    End If
    frmGXBj.lblLcUid.Caption = FMXC.txtXYwy.ToolTipText
    frmGXBj.lblLcRen.Caption = FMXC.txtXYwy.Text

End If
End Sub


Private Sub cmdW6_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If Val(cmdW6.ToolTipText) = 0 And (txtYwy.ToolTipText = mod1.DHid Or mod1.DName = "" Or mod1.DName = "�ܴ���") And txtHtbh.Text = "HMNEW" Then
If mod1.DName <> txtYwy.Text Or lblLc.Caption > 1 Then
Exit Sub
End If
ii = MsgBox("�Ƿ��½���Ʒѯ�۵�?", vbInformation + vbYesNo, "Hello!")
If ii = vbNo Then Exit Sub
    frmGXBj.Visible = False
    tt = "select jzpb,pbid from bjxt_jzpb"
    frmGXBj.adoPb.Close
    frmGXBj.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If IsNull(frmGXBj.adoPb.RecordCount) = True Then
        MsgBox ("��ȡ���ݳ���!")
        Exit Sub
    End If
    Set frmGXBj.comJzpb.RowSource = frmGXBj.adoPb
    frmGXBj.comJzpb.ListField = "jzpb"
    frmGXBj.comJzpb.BoundColumn = "pbid"
    
    
    frmGXBj.Visible = False
    Call modBJD.BJDGXQing
    Call modBJD.gxbjUnLocked
    
'    Set mod1.cmd = createobject("adodb.command")
'    mod1.cmd.ActiveConnection = mod1.CC
'    mod1.cmd.CommandText = "xunJiaAddHT"
'    mod1.cmd.CommandType = adCmdStoredProc
'    mod1.cmd.Parameters("@ywy") = mod1.DName
'    mod1.cmd.Parameters("@uid") = mod1.DHid
'    mod1.cmd.Parameters("@Lx") = 0
'    mod1.cmd.Parameters("@zl") = "����"
'    mod1.cmd.Parameters("@Lcou") = Right(frmGxBiao.cmdCreat.ToolTipText, 1) '��������
'    mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
'    mod1.cmd.Parameters("@lcRen") = mod1.DName
'    mod1.cmd.Parameters("@lcUid") = mod1.DHid
'    mod1.cmd.Parameters("@nLb") = 43
'    mod1.cmd.Parameters("@xmmc") = txtXMMC.Text
'    mod1.cmd.Parameters("@xid") = txtXMMC.ToolTipText
'    mod1.cmd.Parameters("@errch") = ""
'
'    mod1.cmd.Execute


    
timZm = 3 '�½�ѯ�۵�
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "�½�ѯ�۵�"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "��Ʒ"
    mod1.cmd.Parameters("@mt2") = txtXmmc.Text
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mm1") = 43 'NLBֵ
    mod1.cmd.Parameters("@mm2") = txtXmmc.ToolTipText '��Ŀ���
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    LLXX = False
    mod1.cmd.Parameters("@mb1") = 0 'LXֵ
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        If timZm = 3 Then '����
            cmdW6.Enabled = False
        End If
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If

    cmdW6.Enabled = False
Set mod1.cmd = Nothing
    

Else '�����ѯ�۵�
    If Val(cmdW6.ToolTipText) = 0 Then Exit Sub
    Call modBJD.BJDGXQing
    Call modBJD.BJDBound(Val(cmdW6.ToolTipText), "��Ʒ")
    frmGXBj.lblSDJE.Caption = Val(txtD6.Text)
    Call frmGXBj.SDJE(Val(txtD6.Text)) '��̯�ٴ���
    Call modBJD.gxbjLocked
    frmGXBj.optW.Value = True
    
    mod1.BTZ = 36
    frmWait.Visible = False
    frmGXBj.Visible = True
    frmGXBj.ZOrder 0
    frmGXBj.cmdMod.Enabled = True
    frmGXBj.cmdSave.Enabled = False
    frmGXBj.lblLcUid.Caption = FMXC.txtXYwy.ToolTipText
    frmGXBj.lblLcRen.Caption = FMXC.txtXYwy.Text
End If
End Sub

Private Sub cmdWb_Click()
Dim tt As String
On Error Resume Next
Dim Kid As Long
Dim xid As Long
FmxcFK.Visible = False
    'dtgKH.Col = 2
    xid = Me.txtXmmc.ToolTipText
    


    frmWait.Show
    frmWait.ZOrder 0
    
    frmWait.Refresh
    frmWait.faWait.Play
    


    
    Me.Enabled = False
    wbDN.Visible = False
    Me.MousePointer = 11
    mod1.BTZ = 1
    Call mod1.xmQing
    Call mod1.khQing
    Call mod1.xmBound(xid)
    wbDN.lblKid.Caption = wbDN.lblYz.Tag
    Call mod1.khBound(wbDN.lblYz.Tag, "yz")

    wbDN.frmJE.Visible = False

    wbDN.Left = 0
    wbDN.Top = 0
    wbDN.cmdMod.Enabled = False
    wbDN.cmdSave.Enabled = False
    Me.MousePointer = 0
    wbDN.tabKh.Tab = 0

    wbDN.tabKh.TabEnabled(2) = True
    wbDN.tabKh.TabEnabled(0) = True
    

    

    wbDN.modFi = False

    Me.MousePointer = 0
    wbDN.cmdSave.Enabled = False
    wbDN.tabKh.Enabled = True

    wbDN.khAdd = False
    '����Ŀ��,Ĭ�ϵĴ򿪿ͻ�Ϊ��Ŀ����
    wbDN.optYz.Value = True
    wbDN.frmGL.Visible = False
    frmWait.Visible = False
    wbDN.Visible = True
    wbDN.cmdQing.Enabled = False
    wbDN.cmdNew.Enabled = False
    wbDN.cmdRadd.Enabled = False
    wbDN.cmdRdel.Enabled = False
    If wbDN.comXyxz.Text = "��ҵ��˾" Then
        wbDN.frmGL.Visible = True
    End If
    
    '���¶�̬ǩ�ְ�ť�ĳ�ʼ����
        For oo = 1 To 10
           wbDN.lblQM(oo).Left = wbDN.lblQM(oo - 1).Left + 1100
           wbDN.cmdQm(oo).Left = wbDN.cmdQm(oo - 1).Left + 1100
           wbDN.lblTm(oo).Left = wbDN.lblTm(oo - 1).Left + 1100
           mod1.HTP.MoveNext
        Next
End Sub

Private Sub cmdYadd_Click()
Dim tt As String
Dim YYY As Long
Dim hg As Single
Dim oo As Integer
On Error Resume Next
If Val(txtFED.Text) > 100 Then
    MsgBox "��Ȳ��ܳ���100%"
    Exit Sub
End If
If (Val(txtFED.Text) = 0 Or Val(txtYingFu.Text) = 0) And mod1.DName <> "������" Then
Exit Sub
End If

MMdtgYJ.Col = 2
MMdtgYJ.Row = 1
YYY = 0
For oo = 1 To MMdtgYJ.Rows '�������޸�
    YYY = YYY + Val(MMdtgYJ.Text)
Next
YYY = YYY + Val(txtYingFu.Text)
'''''If YYY > Val(txtYj1.Text) Then
'''''    MsgBox "�����޶������ӣ�"
'''''    Exit Sub
'''''End If
tt = "select yjff from htping where htbh='" & txtHtbh.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.HTP.RecordCount) Or mod1.HTP.RecordCount = 0 Then
    MsgBox ("��ȡ���ݴ���1!")
    Exit Sub
End If
If mod1.HTP.Fields("yjff").Value = True Then
    MsgBox ("�����Ѿ�ȫ��֧��,�����ٸ���!")
    Exit Sub
End If


timZm = 13 '��ӽ���
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "����༭"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text) '��ͬ���
    mod1.cmd.Parameters("@mt2") = Trim(txtXmmc.Text) '��Ŀ����
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mm1") = Val(txtFED.Text) / 100 '���
    mod1.cmd.Parameters("@mm2") = Val(txtYingFu.Text) 'Ӧ��
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 1 '��ӽ���
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        If timZm = 3 Then '����

        End If
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


Set mod1.cmd = Nothing

''''''Exit Sub
''''''
''''''Set mod1.cmd = createobject("adodb.command")
''''''mod1.cmd.ActiveConnection = mod1.CC
''''''mod1.cmd.CommandText = "htyjAdd"
''''''mod1.cmd.CommandType = adCmdStoredProc
''''''mod1.cmd.Parameters("@htbh") = Trim(txtHtbh.Text)
''''''mod1.cmd.Parameters("@YED") = Val(txtFED.Text) / 100
''''''mod1.cmd.Parameters("@yingFu") = Val(txtYingFu.Text)
''''''mod1.cmd.Parameters("@xmmc") = Trim(txtXMMC.Text)
''''''mod1.cmd.Execute
''''''Set cmd = Nothing
''''''mod1.mYj.Requery
''''''Set MMdtgYJ.DataSource = mod1.mYj
''''''
''''''Hg = 0
''''''If mod1.mYj.RecordCount > 0 Then
''''''    mod1.mYj.MoveFirst
''''''    Do While Not mod1.mYj.EOF
''''''       Hg = Hg + mod1.mYj.Fields("֧�����").Value
''''''       mod1.mYj.MoveNext
''''''    Loop
''''''End If
'''''''HG = HG + Val(txtYingFu.Text)
'''''''If HG > Val(txtYj.Text) Then
'''''''    MsgBox "��д�������!"
'''''''    txtYingFu.Text = ""
'''''''    Exit Sub
'''''''End If
'''''''End If
''''''txtYj1.Text = Hg
''''''txtLr1.Text = Val(txtJlr1.Text) - Val(txtYj1.Text)
''''''tt = "update htping set yj=" & Val(txtYj1.Text) & ",xmlr=" & Val(txtLr1.Text) & " where htbh='" & txtHtbh.Text & "'"
''''''Set mod1.HTP = CreateObject("adodb.recordset")
''''''mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
End Sub

Private Sub cmdYdel_Click()
Dim tt As String
Dim hg As Single
Dim ii As Integer
Dim Yid As Long
Dim Lc As String
On Error Resume Next
MMdtgYJ.Col = 4
Lc = Val(MMdtgYJ.Text)
MMdtgYJ.Col = 3
Yid = 0
Yid = MMdtgYJ.Text


If Yid = 0 Then
Exit Sub
End If

If Lc > 1 Then
    MsgBox "�˵��Ѿ�����,����ɾ��! ���ȷ��Ҫɾ��,������������ϵ!"
    Exit Sub
End If


ii = MsgBox("�Ƿ�ɾ���˼�¼?", vbQuestion + vbYesNo, "ѯ��")
If ii = vbNo Then
    Exit Sub
End If

timZm = 13 '����༭
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "����༭"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text) '��ͬ���
    mod1.cmd.Parameters("@mt2") = Trim(txtXmmc.Text) '��Ŀ����
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mm1") = Yid
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 0 '����ɾ��
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


Set mod1.cmd = Nothing



Exit Sub




tt = "delete from yongjin where yid=" & Yid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
mod1.mYj.Requery
Set MMdtgYJ.DataSource = mod1.mYj

hg = 0
If mod1.mYj.RecordCount > 0 Then
    mod1.mYj.MoveFirst
    Do While Not mod1.mYj.EOF
       hg = hg + mod1.mYj.Fields("֧�����").Value
       mod1.mYj.MoveNext
    Loop
End If

txtYj1.Text = hg
txtLr1.Text = Val(txtJlr1.Text) - Val(txtYj1.Text)
tt = "update htping set yj=" & Val(txtYj1.Text) & ",xmlr=" & Val(txtLr1.Text) & " where htbh='" & txtHtbh.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
End Sub




Private Sub cmdYview_Click()
Dim tt As String
Dim hg As Single
Dim ii As Integer
Dim Yid As Long
Dim Ywy As String
Dim oo As Integer
On Error Resume Next
MMdtgYJ.Col = 4
Ywy = MMdtgYJ.Text
MMdtgYJ.Col = 3
Yid = 0
Yid = Val(MMdtgYJ.Text)


If Yid = 0 Then
Exit Sub
End If

    
        Dim QFF As Boolean
        mod1.BTZ = 23
        
        frmYjBx.Visible = False
        Call frmYjBx.yjBXQing
        tt = "select * from newYjHt where yid=" & Yid
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        frmYjBx.lblQy.Caption = mod1.HTP.Fields("qy").Value
        frmYjBx.lblBM.Caption = mod1.HTP.Fields("bm").Value
        frmYjBx.lblXmmc.Caption = mod1.HTP.Fields("��Ŀ����").Value
        frmYjBx.lblHtbh.Text = mod1.HTP.Fields("��ͬ���").Value
        frmYjBx.lblHtze.Caption = mod1.HTP.Fields("��ͬ���").Value
        frmYjBx.lblYf.Caption = mod1.HTP.Fields("Ӧ��").Value
        frmYjBx.lblED.Caption = mod1.HTP.Fields("�տ���").Value
        frmYjBx.lblYid.Caption = mod1.HTP.Fields("yid").Value
        frmYjBx.lblYwy.Caption = mod1.HTP.Fields("������").Value
        frmYjBx.lblUid.Caption = mod1.HTP.Fields("uid").Value
        frmYjBx.lblLcRen.Caption = mod1.HTP.Fields("lcren").Value
        frmYjBx.lblLcUid.Caption = mod1.HTP.Fields("lcuid").Value
        frmYjBx.lblLc.Caption = mod1.HTP.Fields("lc").Value
        frmYjBx.lblFwid.Caption = mod1.HTP.Fields("fwid").Value
        frmYjBx.txtCXF.Text = mod1.HTP.Fields("cxf").Value
        frmYjBx.txtBz.Text = mod1.HTP.Fields("��ע").Value
        Pwf = mod1.HTP.Fields("pwf").Value
        QFF = mod1.HTP.Fields("Qff").Value
        tt = "select yj from htping where htbh='" & frmYjBx.lblHtbh.Text & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        frmYjBx.lblYj.Caption = mod1.HTP.Fields("yj").Value

        tt = "select sum(Ӧ��)+sum(cxf) from newyjht where ��ͬ���='" & frmYjBx.lblHtbh.Text & "' and ֧����=1"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        '���÷�������е�����֧��
        'ʵ�ʱ�
        tt = "Select sum(zFu) as zfu from yjz where htbh='" & frmYjBx.lblHtbh.Text & "'"
        mod1.HTT.Close
        mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText

        If IsNull(mod1.HTP.Fields(0).Value) = True Then
            Ny = 0
        Else
            Ny = mod1.HTP.Fields(0).Value
        End If
        frmYjBx.lblCf.Caption = Ny + mod1.HTT.Fields("zfu").Value
'        If IsNull(mod1.HTP.Fields(0).Value) = True Then
'            frmYjBx.lblCf.Caption = 0
'        Else
'            frmYjBx.lblCf.Caption = mod1.HTP.Fields(0).Value
'        End If
Call frmYjBx.Lren(Val(lblMHid.Caption))
        
'''''        For oo = 0 To 6
'''''            frmYjBx.lblTm(oo).Caption = ""
'''''            frmYjBx.cmdQm(oo).Caption = ""
'''''            frmYjBx.lblQM(oo).Visible = False
'''''            frmYjBx.lblTm(oo).Visible = False
'''''            frmYjBx.cmdQm(oo).Visible = False
'''''        Next
''''''''        '�򿪰�ť
''''''''        tt = "select * from qmrz where btz=23 and qdbh='" & TBh & "' order by zid"
''''''''        Set mod1.HTP = CreateObject("adodb.recordset")
''''''''        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''''
''''''''        mod1.HTP.MoveFirst
''''''''
''''''''        For oo = 0 To 6
''''''''            frmYjBx.lblQM(oo).Caption = mod1.HTP.Fields("qLabel").Value
''''''''            If mod1.HTP.Fields("xf").Value = True Then
''''''''                frmYjBx.cmdQm(oo).Caption = mod1.HTP.Fields("qren").Value
''''''''                If frmYjBx.cmdQm(oo).Caption = "�Ͼ��쾭��" Then
''''''''                    frmYjBx.cmdQm(oo).Caption = "�Ͼ��쾭��"
''''''''                End If
''''''''                frmYjBx.lblTm(oo).Caption = mod1.HTP.Fields("qrq").Value
''''''''            End If
''''''''            frmYjBx.cmdQm(oo).Visible = True
''''''''            frmYjBx.lblQM(oo).Visible = True
''''''''            frmYjBx.lblTm(oo).Visible = True
''''''''            mod1.HTP.MoveNext
''''''''        Next
        
'�ж�����ǩ�ְ�ť,��û��,�����
'''''If frmYjBx.lblYwy.Caption <> "" Then
'''''    tt = "select * from qmrz where btz=23 and qdbh='" & Yid & "' order by zid"
'''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''
'''''    mod1.HTP.MoveFirst
'''''    For oo = 0 To 6
'''''        frmYjBx.lblQM(oo).Caption = mod1.HTP.Fields("qLabel").Value
'''''        If mod1.HTP.Fields("xf").Value = True Then
'''''            frmYjBx.cmdQm(oo).Caption = mod1.HTP.Fields("qren").Value
'''''            If frmYjBx.cmdQm(oo).Caption = "�Ͼ��쾭��" Then
'''''                frmYjBx.cmdQm(oo).Caption = "�Ͼ��쾭��"
'''''            End If
'''''            frmYjBx.lblTm(oo).Caption = mod1.HTP.Fields("qrq").Value
'''''        End If
'''''        frmYjBx.cmdQm(oo).Visible = True
'''''        frmYjBx.lblQM(oo).Visible = True
'''''        frmYjBx.lblTm(oo).Visible = True
'''''        mod1.HTP.MoveNext
'''''    Next
'''''    If frmYjBx.lblQM(5).Caption = "��֧��" Then
'''''        frmYjBx.lblQM(6).Visible = False
'''''        frmYjBx.cmdQm(6).Visible = False
'''''        frmYjBx.lblTm(6).Visible = False
'''''    End If
'''''    If Pwf = True And frmYjBx.cmdQm(5).Caption = "" And frmYjBx.cmdQm(6).Visible = False Then '��֧����ʾ
'''''        frmYjBx.cmdQm(5).Caption = frmYjBx.cmdQm(2).Caption
'''''        frmYjBx.lblTm(5).Caption = frmYjBx.lblTm(4).Caption
'''''    End If
'''''
'''''Else
'''''
'''''End If
        
        
        

        If QFF = False And mod1.DName = "�Ǽ���" And frmYjBx.lblLc.Caption = 7 Then
            frmYjBx.cmdWb.Visible = True
        Else
            frmYjBx.cmdWb.Visible = False
        End If
        
        frmYjBx.lblLcRen.Caption = mod1.DName
        frmYjBx.lblLcUid.Caption = mod1.DHid
'''''        If frmYjBx.lblQM(6).Caption = "" Or frmYjBx.lblQM(5).Caption = frmYjBx.lblQM(6).Caption Then
'''''            frmYjBx.lblQM(6).Visible = False
'''''            frmYjBx.cmdQm(6).Visible = False
'''''            frmYjBx.lblTm(6).Visible = False
'''''        End If
        
                    '������֤
        On Error GoTo YZERR9
        tt = "insert into HMText.dbo.ML (NB,NBLX,trq,bh,ywy,uid,Bz,mt3) values ('����','�鿴',getdate(),'" & frmYjBx.lblYid.Caption & _
            "','" & mod1.DName & "','" & mod1.DHid & "' ,'" & frmYjBx.lblXmmc.Caption & "','" & frmYjBx.lblHtbh.Text & "')"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Set mod1.HTP = Nothing
        frmYjBx.OptT1.Value = False: frmYjBx.optT2.Value = False
        frmYjBx.Show
        Exit Sub
YZERR9:
        MsgBox "������ϣ�������һ�Σ�������������"
        Exit Sub


End Sub


Private Sub comYjRen_Click()
comYjRen.ToolTipText = Rid(comYjRen.ListIndex)
End Sub


Private Sub dt3_CloseUp()
txtF.Text = dt3.Value
End Sub


Private Sub dt4_CloseUp()
txtL.Text = dt4.Value
End Sub


Private Sub Form_DblClick()
Dim ii As Integer
Dim tt As String
Dim Je1 As Single, Je2 As Single, Je3 As Single, Je4 As Single, Je5 As Single, Je6 As Single, Je7 As Single
Dim Bid1 As Long, Bid6 As Long, Bid7 As Long
Dim Ra
If mod1.DName = "������" Or mod1.DName = "�ռ���" Or mod1.DName = "������" Then
    frmYj.Visible = True
    Exit Sub
End If
'If mod1.DName <> "���ӱ" Then Exit Sub
If mod1.DName <> "�Ǽ���" Or mod1.DName <> "����" Or mod1.DName = "�Ǽ���" Then Exit Sub
dtgFL.Col = 2: dtgFL.Row = 1
Je1 = Val(dtgFL.Text)
dtgFL.Col = 4: Bid1 = Val(dtgFL.Text)
dtgFL.Col = 2: dtgFL.Row = 2
Je2 = Val(dtgFL.Text)
dtgFL.Row = 3: Je3 = Val(dtgFL.Text)
dtgFL.Row = 4: Je4 = Val(dtgFL.Text)
dtgFL.Row = 5: Je5 = Val(dtgFL.Text)
dtgFL.Row = 6: Je6 = Val(dtgFL.Text): dtgFL.Col = 4: Bid6 = Val(dtgFL.Text): dtgFL.Col = 2
dtgFL.Row = 7: Je7 = Val(dtgFL.Text): dtgFL.Col = 4: Bid7 = Val(dtgFL.Text)

tt = "select sum(amount) from SD30301_�º�������.dbo.s_order where billcode like '%" & txtHtbh.Text & "%' and billstate=1 and closed=0"
Set mod1.HTP = CreateObject("adodb.recordset")
On Error GoTo ZXERR
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
'If Round(Ra(0, 0), 1) = Val(txtHtze.Text) Then
If Val(txtHtze.Text) > 0 Then ' ��ʱ���Ժ�����Ϣ����ٴ�Ķ���
    timZm = 19 'ִ��֪ͨ
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "ִ��֪ͨ"
    mod1.cmd.Parameters("@bh") = Val(lblMHid.Caption)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtXmmc.Text
    mod1.cmd.Parameters("@mt2") = txtHtbh.Text
    mod1.cmd.Parameters("@mt3") = comQy.Text
    mod1.cmd.Parameters("@mt4") = txtXYwy.Text
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Je1
    mod1.cmd.Parameters("@mm2") = Je2
    mod1.cmd.Parameters("@mm3") = Je3
    mod1.cmd.Parameters("@mm4") = Je4
    mod1.cmd.Parameters("@mm5") = Je5
    mod1.cmd.Parameters("@mm6") = Je6
    mod1.cmd.Parameters("@mm7") = Je7
    mod1.cmd.Parameters("@mm8") = Bid1
    mod1.cmd.Parameters("@mm9") = Bid6
    mod1.cmd.Parameters("@mm10") = Bid7
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        If timZm = 2 Then '����
            cmdSave.Enabled = False
        End If
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
        frmFk.Visible = False
        frmFX.Visible = False
        
    End If

    
Set mod1.cmd = Nothing
Else
    ii = MsgBox("�ٴ����۶�����˿�������©,Ŀǰ�ٴ�������Ϊ" & Ra(0, 0) & ",������Ϣ���Ϊ:" & txtHtze.Text, vbInformation, "������ٴ�!")
    Exit Sub
End If
Exit Sub
ZXERR:
MsgBox "����!"
End Sub

Private Sub Label6_Click()
FmxcFK.Show
FmxcFK.ZOrder 0
FmxcFK.Enabled = True
End Sub

Private Sub Label9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mod1.Kyj = True And Button = 2 Then

       ' tt=inputbox(""
        timYj.Enabled = True

End If
End Sub




Private Sub dtgFL_DblClick()
Dim tt As String
Dim ii As Integer
Dim Bid As Long
On Error Resume Next
FmxcFK.Visible = False
dtgFL.Col = 4
Bid = Mid(Trim(dtgFL.Text), 4, Len(Trim(dtgFL.Text)) - 3)
If Bid > 0 Then
    mod1.BTZ = 36
            If mod1.Mname = "������" Or mod1.DName = "��Ʒ¼��Ա" Or mod1.DName = "������" Then
                Call frmGxbjNew.Initialize
                Call frmGxbjNew.Bound(Bid)
                mod1.BTZ = 36
                frmWait.Visible = False
                frmGxbjNew.Visible = True
                frmGxbjNew.ZOrder 0
                frmGxbjNew.cmdMod.Enabled = True
                frmGxbjNew.cmdSave.Enabled = False
                Exit Sub
            End If
        If dtgFL.Row > 0 And dtgFL.Row < 6 Then
            Call frmWBXX.Qing
            Call frmWBXX.Bound(Bid)
            'Call frmWBXNew.Bound(Val(dtgFL.Text))
            frmWBXX.Show
            frmWBXX.ZOrder 0
            Exit Sub
        Else
            If mod1.Mname = "������" Then
                Call frmGxbjNew.Initialize
                frmGxbjNew.Show
                frmGxbjNew.lblTitle.Caption = "<<=��ѡ���ѯ��,����ѡ��ֱ������ԭ�����!"
            Else
                Call modBJD.BJDGXQing
                If dtgFL.Row = 6 Then
                    Call modBJD.BJDBound(Bid, "���")
                    Call frmGXBj.SDJE(Val(txtD5.Text)) '��̯�ٴ���
                ElseIf dtgFL.Row = 7 Then
                    Call modBJD.BJDBound(Bid, "��Ʒ")
                    Call frmGXBj.SDJE(Val(txtD6.Text)) '��̯�ٴ���
                End If
                Call frmGXBj.dtgMaFF
    
                Call modBJD.gxbjLocked
                frmGXBj.optW.Value = True
                mod1.BTZ = 36
                frmWait.Visible = False
                frmGXBj.Visible = True
                frmGXBj.ZOrder 0
                frmGXBj.cmdMod.Enabled = True
                frmGXBj.cmdSave.Enabled = False
                frmGXBj.frmJ.Visible = True
    
    '''            frmGXBj.lblLcUid.Caption = FMXC.txtXYwy.ToolTipText
    '''            frmGXBj.lblLcRen.Caption = FMXC.txtXYwy.Text
            End If
        End If
        Exit Sub
End If

If txtHtbh.Text <> "HMNEW" Then
    Exit Sub
End If
If Bid = 0 And (txtYwy.ToolTipText = mod1.DHid Or txtXYwy.ToolTipText = mod1.DHid Or mod1.DName = "" Or mod1.DName = "�ܴ���" Or mod1.DName = "������") Then
'''''''    If mod1.DName <> txtYwy.Text Or lblLc.Caption > 1 Then
'''''''    Exit Sub
'''''''    End If
    If dtgFL.Row = 6 Then
    End If
    ii = MsgBox("�Ƿ��½�ѯ�۵�?", vbInformation + vbYesNo, "Hello!")
'''    MsgBox ("���ڲ����У�����϶����ã����½⣡")
'''    Exit Sub
    If ii = vbNo Then Exit Sub
   
    
'''''    frmWBXJ.Visible = False
'''''    Call modBJD.BJDWBQing
'''''    Call modBJD.wbxjUnLocked
    
    
timZm = 3 '�½�ѯ�۵�
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "�½�ѯ�۵�"
    mod1.cmd.Parameters("@bh") = lblMHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "�˹�"
    LLXX = True
    If dtgFL.Row = 6 Then
        mod1.cmd.Parameters("@mt1") = "���"
        LLXX = False
    ElseIf dtgFL.Row = 7 Then
        mod1.cmd.Parameters("@mt1") = "��Ʒ"
        LLXX = False
    ElseIf dtgFL.Row = 3 Or dtgFL.Row = 4 Or dtgFL.Row = 5 Then
        mod1.cmd.Parameters("@mt1") = "�ְ�"
        LLXX = True
    End If
    mod1.cmd.Parameters("@mt2") = txtXmmc.Text
    mod1.cmd.Parameters("@mt3") = txtADR.Text
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
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
    mod1.cmd.Parameters("@mm1") = 88 'NLBֵ
    mod1.cmd.Parameters("@mm2") = txtXmmc.ToolTipText '��Ŀ���
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 0
   'Exit Sub
    mod1.cmd.Parameters("@mm11") = 0
    mod1.cmd.Parameters("@mm12") = 0
    mod1.cmd.Parameters("@mm13") = 0
    mod1.cmd.Parameters("@mm14") = 0
    mod1.cmd.Parameters("@mm15") = 0
    mod1.cmd.Parameters("@mm16") = 0
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0

    mod1.cmd.Parameters("@mb1") = 1 'LXֵ
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
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"

        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
    
    mod1.BTZ = 36
End If
End Sub


Private Sub frmYj_Click()
dtgSD.Visible = False
End Sub

Private Sub Label1_DblClick()
Dim tt As String
On Error Resume Next
        mod1.BTZ = 1
        tt = "Select xid,kid from khren where rid=" & Val(comYjRen.ToolTipText)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        
'''''        If mod1.DKZ(mod1.HTP.Fields("xid").Value, 6) = True Then
'''''        MsgBox "��ݱ�����" & mod1.DKRen & "��,���Ժ�����,������������ϵ."
'''''        Dialog.Enabled = True
'''''        Exit Sub
'''''        End If

          wbDN.Visible = False
          Me.MousePointer = 11
'''''          '��¼����־
'''''          Call mod1.zhuDa(3, mod1.HTP.Fields("xid").Value)
          Call mod1.xmQing
          Call mod1.khQing
          
          Call mod1.khFuBound(mod1.HTP.Fields("kid").Value, mod1.HTP.Fields("xid").Value, Val(comYjRen.ToolTipText))
        
          wbDN.cmdMod.Enabled = False
          wbDN.cmdSave.Enabled = False
          wbDN.tabKh.Tab = 1
'          wbDN.cmdRadd.Enabled = False
'          wbDN.cmdNew.Enabled = False
          wbDN.khAdd = False
          frmWait.Visible = False
          wbDN.Visible = True
          'wbDN.adoRen.Recordset.Move 0
          Me.MousePointer = 0
          If wbDN.lblYwy.Caption = mod1.DName Or wbDN.lblXywy.Caption = mod1.DName Then
              wbDN.cmdMod.Enabled = True
          Else
              wbDN.cmdMod.Enabled = False
          End If
          wbDN.lblLcRen.Caption = mod1.DName
          wbDN.lblLcUid.Caption = mod1.DHid
          wbDN.cmdMod.Enabled = True
End Sub


Private Sub lblCBZE_DblClick()
Dim tt As String
Dim oo As Integer
Dim Ra
Dim La
FmxcZBR.dtgZBr.Clear: FmxcZBR.dtgN.Clear
FmxcZBR.dtgFF
FmxcZBR.Show
FmxcZBR.ZOrder 0
tt = "select bh,gui,ze,zid from htZui where hid=" & Me.lblMHid.Caption & " order by zid"
tt = "SELECT dbo.htZui.Bh, dbo.htZui.Gui, SUM(dbo.htZuiDetail.Ze) AS Ze, dbo.htZui.Zid FROM dbo.htZui LEFT OUTER JOIN dbo.htZuiDetail ON dbo.htZui.Zid = dbo.htZuiDetail.Zid" & _
    " where dbo.htzui.hid=" & Me.lblMHid.Caption & " and htzui.delf=1 GROUP BY dbo.htZui.Bh, dbo.htZui.Gui, dbo.htZui.Zid order by dbo.htzui.zid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.BOF = True Then Exit Sub
Ra = mod1.HTP.GetRows
La = UBound(Ra, 2) + 1
Set mod1.HTP = Nothing
Call FmxcZBR.dtgFF
On Error Resume Next
For oo = 1 To La
    FmxcZBR.dtgZBr.Row = oo
    FmxcZBR.dtgZBr.Col = 0: FmxcZBR.dtgZBr.Text = Ra(0, oo - 1)
    FmxcZBR.dtgZBr.Col = 1: FmxcZBR.dtgZBr.Text = Ra(1, oo - 1)
    FmxcZBR.dtgZBr.Col = 2: FmxcZBR.dtgZBr.Text = Ra(2, oo - 1)
    FmxcZBR.dtgZBr.Col = 3: FmxcZBR.dtgZBr.Text = Ra(3, oo - 1)
    
    FmxcZBR.dtgN.Row = oo
    FmxcZBR.dtgN.Col = 0: FmxcZBR.dtgN.Text = Ra(0, oo - 1)
    FmxcZBR.dtgN.Col = 1: FmxcZBR.dtgN.Text = Ra(1, oo - 1)
    FmxcZBR.dtgN.Col = 2: FmxcZBR.dtgN.Text = Ra(2, oo - 1)
    FmxcZBR.dtgN.Col = 3: FmxcZBR.dtgN.Text = Ra(3, oo - 1)
Next

End Sub

Private Sub MMdtgBao_DblClick()
Dim tt As String
 '"select Sl AS ����,dj AS �ɱ�����,Wdj AS �������,jdj AS ��׼����, Whg AS ����ϼ�,jhg AS ��׼�ϼ� from xunjiamxView where bid=@bid5;
On Error Resume Next
MMdtgBao.Col = 11
txtTl.Text = MMdtgBao.Text
MMdtgBao.Col = 12
txtDj.Text = Val(MMdtgBao.Text)
MMdtgBao.Col = 16
liD = Val(MMdtgBao.Text)
MMdtgBao.Col = 17
LLid = Val(MMdtgBao.Text)
Set MMdtgMa.DataSource = Nothing
MMdtgMa.Refresh
tt = "select * from xunJiaMxView where lid=" & liD
mod1.mGx.Close
mod1.mGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.mGx.RecordCount = 1 Then
    Set MMdtgMa.DataSource = mod1.mGx
    MMdtgMa.Visible = True
Else
    MMdtgMa.Visible = False
End If
End Sub

Private Sub mmdtgcp_Click()
Dim tt As String
Dim liD As Long
On Error Resume Next
MMdtgCP.Col = 11
txtCL.Text = MMdtgCP.Text
MMdtgCP.Col = 12
txtCj.Text = MMdtgCP.Text
MMdtgCP.Col = 16
liD = MMdtgCP.Text
tt = "select * from xunJiaMxView where lid=" & liD
mod1.mGxCP.Close
mod1.mGxCP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set MMdtgCPCB.DataSource = mod1.mGxCP
End Sub

Private Sub mmdtgcp_RowColChange()
Dim tt As String
Dim liD As Long
On Error Resume Next
MMdtgCP.Col = 11
txtCL.Text = MMdtgCP.Text
MMdtgCP.Col = 12
txtCj.Text = MMdtgCP.Text
MMdtgCP.Col = 16
liD = MMdtgCP.Text
tt = "select * from xunJiaMxView where lid=" & liD
mod1.mGxCP.Close
mod1.mGxCP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set MMdtgCPCB.DataSource = mod1.mGxCP
End Sub


Private Sub mmdtgfk_Click()
On Error Resume Next
If Val(MMdtgFk.Text) = 0 Then Exit Sub
MMdtgFk.Col = 1
dtpYf.Value = MMdtgFk.Text
txtYrq.Text = MMdtgFk.Text
MMdtgFk.Col = 2
txtYed.Text = Val(MMdtgFk.Text)
MMdtgFk.Col = 3
txtYje.Text = Val(MMdtgFk.Text)
MMdtgFk.Col = 4
lblFid.Caption = MMdtgFk.Text
End Sub

Private Sub mmdtgfk_RowColChange()
On Error Resume Next
If Val(MMdtgFk.Text) = 0 Then Exit Sub
MMdtgFk.Col = 1
dtpYf.Value = MMdtgFk.Text
txtYrq.Text = MMdtgFk.Text
MMdtgFk.Col = 2
txtYed.Text = Val(MMdtgFk.Text)
MMdtgFk.Col = 3
txtYje.Text = Val(MMdtgFk.Text)
MMdtgFk.Col = 4
lblFid.Caption = MMdtgFk.Text
End Sub


Private Sub dtpYf_CloseUp()
txtYrq.Text = dtpYf.Value
End Sub

Private Sub Form_Click()

frmQm.Visible = False
lblTX.Visible = False
Me.FO = 1

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 76 Then
'''    If mod1.Kyj = True Then
'''        If frmYj.Visible = False Then
'''            frmYj.Visible = True
'''            lblTcBe.Visible = True
'''            txtTcBe.Visible = True
'''        Else
            frmYj.Visible = False
            lblTcBe.Visible = False
            txtTcBe.Visible = False
'''        End If
'''   End If
'''
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
dtgFL.Left = 0
dtgFL.Top = 240
Call Me.FLGG

dtgSD.Row = 20
dtgSD.ColWidth(0) = 2490
dtgSD.Top = 5460
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0

frmJi.BorderStyle = 0

'''''''''Set mWb = CreateObject("adodb.recordset")
'''''''''Set mLj = CreateObject("adodb.recordset")
''''''''''Set adoOid = CreateObject("adodb.recordset")
'''''''''Set mod1.mBx = CreateObject("adodb.recordset")
'''''''''Set mod1.mGx = CreateObject("adodb.recordset")
'''''''''Set mod1.mFk = CreateObject("adodb.recordset")
'''''''''Set mod1.mYj = CreateObject("adodb.recordset")
'''''''''Set mod1.mBxCP = CreateObject("adodb.recordset")
'''''''''Set mod1.mGxCP = CreateObject("adodb.recordset")
'''''''''
'''''''''Set mod1.mA = CreateObject("adodb.recordset")
'''''''''Set mod1.mB = CreateObject("adodb.recordset")

MMdtgMa.ColWidth(0) = 300


MMdtgBao.ColWidth(0) = 300
'''MMdtgBao.ColWidth(8) = 2000
'''MMdtgBao.ColWidth(15) = 0
'''MMdtgBao.ColWidth(16) = 0

MMdtgCP.ColWidth(0) = 300
'''MMdtgCP.ColWidth(8) = 2000
'''MMdtgCP.ColWidth(15) = 0
'''MMdtgCP.ColWidth(16) = 0

MMdtgCPCB.ColWidth(0) = 300
'''MMdtgCPCB.ColWidth(8) = 2000
'''MMdtgCPCB.ColWidth(13) = 0
'''MMdtgCPCB.ColWidth(15) = 0
'''MMdtgCPCB.ColWidth(18) = 0
'''MMdtgCPCB.ColWidth(19) = 0
'''MMdtgCPCB.ColWidth(20) = 0
'''MMdtgCPCB.ColWidth(22) = 0


MMdtgBao.Left = 0
MMdtgBao.Top = 0
frmYj.BorderStyle = 0


MMdtgA.ColWidth(0) = 300
MMdtgA.ColWidth(2) = 2000
MMdtgA.ColWidth(3) = 700
MMdtgA.ColWidth(4) = 0

MMdtgFk.ColWidth(0) = 300
MMdtgFk.ColWidth(1) = 1300
MMdtgFk.ColWidth(2) = 900
MMdtgFk.ColWidth(4) = 0

MMdtgYJ.ColWidth(0) = 300
MMdtgYJ.ColWidth(3) = 0
MMdtgYJ.ColWidth(4) = 0

frmFk.BorderStyle = 0
frmNb.BorderStyle = 0
frmTime.BorderStyle = 0
dtpYf.Value = mod1.DQda
dt3.Value = mod1.DQda
dt4.Value = mod1.DQda

frmQm.Left = 810
frmQm.Top = 7440
frmQm.Visible = False

chkA.ForeColor = &H80000012
chkB.ForeColor = &H80000012
chkC.ForeColor = &H80000012
chkD.ForeColor = &H80000012
chkE.ForeColor = &H80000012
chkF.ForeColor = &H80000012
dtpJTF.Value = mod1.DQda
dtpQkF.Value = mod1.DQda
dtgJTf.ColWidth(0) = 300
dtgJTf.ColWidth(1) = 2000
dtgJTf.ColWidth(3) = 5000
dtgJTf.ColWidth(4) = 0
dtgQkf.ColWidth(0) = 300
dtgQkf.ColWidth(1) = 2000
dtgQkf.ColWidth(3) = 5000
dtgQkf.ColWidth(4) = 0
dtgyjF.ColWidth(0) = 300
dtgyjF.ColWidth(1) = 2000
dtgyjF.ColWidth(3) = 5000
dtgyjF.ColWidth(4) = 0


End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
If htBrow.Visible = True Then
    htBrow.adoBr.Requery
    Set htBrow.dtgBr.DataSource = htBrow.adoBr
    htBrow.Enabled = True
    htBrow.ZOrder 0
ElseIf htBrowG.Visible = True Then
    htBrowG.Enabled = True
    htBrowG.ZOrder 0
ElseIf Dialog.Enabled = True Then
    Dialog.ZOrder 0
    Dialog.Enabled = True
ElseIf FmxcXB.Visible = True Then
    FmxcXB.Enabled = True
    FmxcXB.ZOrder 0
End If
Cancel = True
End Sub

Private Sub tabGc_Click(PreviousTab As Integer)
'Dim oo As Integer
'For oo = 0 To 5
'frmC(oo).Visible = False
'Next
'frmgc(tabGc.Tab).Visible = True
If tabGc.Tab = 0 Then
    MMdtgBao.Visible = False
Else
    MMdtgBao.Visible = True
End If
End Sub

Private Sub tabHt_Click(PreviousTab As Integer)
frmQm.Visible = False
If tabHt.Tab = 1 Then
    'txtFbnr.Visible = False
    'txtWBNR.Visible = False
    If Val(txtH1.Text) > 0 Then
        tabGc.TabVisible(0) = True
    End If
    If Val(txtH2.Text) > 0 Then
        tabGc.TabVisible(1) = True
    End If
    If Val(txtW3.Text) > 0 Then
        tabGc.TabVisible(4) = True
        'txtFbnr.Visible = True
    End If
    If Val(txtW4.Text) > 0 Then
        tabGc.TabVisible(5) = True
        'txtWBNR.Visible = True
    End If
    If Val(txtH5.Text) > 0 Or Val(txtW5.Text) > 0 Then
        tabGc.TabVisible(2) = True
    End If
    If Val(txtH6.Text) > 0 Or Val(txtW6.Text) > 0 Then
        tabGc.TabVisible(3) = True
    End If
End If

End Sub

Private Sub tabHt_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 76 Then
'''    If mod1.Kyj = True Then
'''        If frmYj.Visible = False Then
'''            frmYj.Visible = True
'''            lblTcBe.Visible = True
'''            txtTcBe.Visible = True
'''        Else
            frmYj.Visible = False
            lblTcBe.Visible = False
            txtTcBe.Visible = False
'''        End If
'''   End If
'''
End If
End Sub


Private Sub tabHt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'MsgBox X
'MsgBox Y

'''If mod1.Kyj = True And Button = 2 Then
'''    If X > 15075 And Y < 135 Then
'''       ' tt=inputbox(""
'''        timYj.Enabled = True
'''    Else
'''        timYj.Enabled = False
'''    End If
'''End If
End Sub

Private Sub tabHt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
timYj.Enabled = False
End Sub

Private Sub timQuit_Timer()
Dim oo As Integer
Dim ii As Integer
On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
Dim tt As String
If timZm = 2 Then '���Ϊ��Ӻ�ͬ����
    Call modNewHT.NewLocked
    cmdSave.Enabled = False
    If Val(lblLc.Caption) = 0 Then
        lblLc.Caption = 1
    End If
ElseIf timZm = 3 Then '�½����ѯ�۵�
    frmGXBj.OPTN.Value = True
    frmGxbjNew.frmSd.Visible = True
ElseIf timZm = 10 Then 'ǩ��
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
    
ElseIf timZm = 11 Then
    cmdHT.Visible = False
    If lblHtxz.Caption = "ά��" Then
        frmDate.Visible = True
    End If
     MsgBox "˫����ͬ���,���Ը��ӵ��Ӻ�ͬ"
ElseIf timZm = 12 Then 'ɾ����ͬ
    Me.Visible = False
    If htBrow.Visible = True Then
        htBrow.Enabled = True
        htBrow.ZOrder 0
        htBrow.adoBr.Requery
        Set htBrow.dtgBr.DataSource = htBrow.adoBr
    ElseIf htBrowG.Visible = True Then
        htBrowG.Enabled = True
        htBrowG.ZOrder 0

    ElseIf Dialog.Visible = True Then
        Dialog.Enabled = True
        Dialog.ZOrder 0
        Call mod1.refEnvent(1)
    End If
ElseIf timZm = 13 Then '��ӽ���
    txtFED.Text = ""
    txtYingFu.Text = ""
    Dim Ra
    Dim ua
    tt = "select yED as �տ���,YingFu as ֧�����,yid from yongjin where htbh='" & txtHtbh.Text & "' order by yid"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    ua = UBound(Ra, 2)
    Set mod1.HTP = Nothing
    MMdtgYJ.Clear
    FMXC.MMdtgYJ.Row = 0: FMXC.MMdtgYJ.Col = 1: FMXC.MMdtgYJ.Text = "�տ���"
    FMXC.MMdtgYJ.Col = 2: FMXC.MMdtgYJ.Text = "֧�����"
    For oo = 1 To ua + 1
        MMdtgYJ.Row = oo
        For ii = 1 To 3
            MMdtgYJ.Col = ii
            MMdtgYJ.Text = Trim(Ra(ii - 1, oo - 1))
        Next
    Next
    
    Dim CB As Double
    FMXC.MMdtgYJ.Row = 0: FMXC.MMdtgYJ.Col = 5: FMXC.MMdtgYJ.Text = "�ο����"
    FMXC.MMdtgYJ.Row = 1
    FMXC.MMdtgYJ.Col = 1
    Do While Not Val(FMXC.MMdtgYJ.Text) = 0
        
        CB = (Val(FMXC.txtHtze.Text) - Val(FMXC.txtCbze1.Text)) * Val(FMXC.MMdtgYJ.Text)
        FMXC.MMdtgYJ.Col = 5
        FMXC.MMdtgYJ.Text = CB
        FMXC.MMdtgYJ.Col = 1
        FMXC.MMdtgYJ.Row = FMXC.MMdtgYJ.Row + 1
        CB = 0
    Loop
ElseIf timZm = 19 Then 'ִ��֪ͨ
    MsgBox "�Ѿ��ɹ�֪ͨ:" & lblTX.Caption & "!"
ElseIf timZm = 20 Then '��ͬ����
    Call FmxcNew.Bound(Val(Me.lblMHid.Caption))
    FmxcNew.Show
    FmxcNew.ZOrder 0
    FMXC.Visible = False
    MsgBox ("�����ɹ�! (���ҵ���������,��������һ�����Ӧ����)")
End If
timQuit.Enabled = False

End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2,mt3 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '�ύ�ɹ�
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 2 Then
        cmdSave.Enabled = False
    ElseIf timZm = 3 Then
        If LLXX = False Then
            If mod1.Mname = "������" Or mod1.DName = "лѩ÷" Then
                Call frmGxbjNew.Initialize
                frmGxbjNew.lblBh.ToolTipText = mod1.WP.Fields("mt2").Value
                frmGxbjNew.lblBh.Caption = "XJD" & mod1.WP.Fields("mt2").Value
                frmGxbjNew.lblLc.Caption = 1
                frmGxbjNew.lblLcRen.Caption = mod1.DName
                frmGxbjNew.lblLcUid.Caption = mod1.DHid
                frmGxbjNew.lblYwy.Caption = mod1.DName
                frmGxbjNew.lblUid.Caption = mod1.DHid
                frmGxbjNew.lblZl.Caption = mod1.WP.Fields("mt1").Value
                If mod1.WP.Fields("mt1").Value = "���" Or mod1.WP.Fields("mt1").Value = "���ѯ�۵�" Then
                    cmdW5.ToolTipText = mod1.WP.Fields("mt2").Value
                    dtgFL.Row = 6: dtgFL.Col = 4: dtgFL.Text = "XJD" & mod1.WP.Fields("mt2").Value
                ElseIf mod1.WP.Fields("mt1").Value = "��Ʒ" Then
                    cmdW6.ToolTipText = mod1.WP.Fields("mt2").Value
                    dtgFL.Row = 7: dtgFL.Col = 4: dtgFL.Text = "XJD" & mod1.WP.Fields("mt2").Value
                ElseIf mod1.WP.Fields("mt1").Value = "�ְ�" Then
                        dtgFL.Row = 3: dtgFL.Col = 4: dtgFL.Text = "XJD" & mod1.WP.Fields("mt2").Value
                End If
                frmGxbjNew.txtXmmc = txtXmmc.Text
                frmGxbjNew.txtXmmc.ToolTipText = txtXmmc.ToolTipText
                frmGxbjNew.txtHg.Locked = True
                frmGxbjNew.cmdHT.ToolTipText = FMXC.lblMHid.Caption
                
                frmGxbjNew.cmdMod.Enabled = False
                
                frmGxbjNew.cmdSave.Enabled = True

                frmGxbjNew.lblZl.ForeColor = &HC000C0

                frmGxbjNew.txtMj.Locked = True
                frmGxbjNew.txtDj.Locked = True
                mod1.BTZ = 36
                frmGxbjNew.Visible = True
                Call frmGxbjNew.initializeForm
                Exit Sub
            End If
            Call modBJD.BJDGXQing
            frmGXBj.lblBid.Caption = mod1.WP.Fields("mt2").Value
            frmGXBj.lblBh.Caption = "XJD" & mod1.WP.Fields("mt2").Value
            frmGXBj.lblLcou.Caption = 3 '��������
            frmGXBj.lblLc.Caption = 1
            frmGXBj.lblLcRen.Caption = mod1.DName
            frmGXBj.lblLcUid.Caption = mod1.DHid
            frmGXBj.lblNlb.Caption = 43
            frmGXBj.lblYwy.Caption = mod1.DName
            frmGXBj.lblUid.Caption = mod1.DHid
            frmGXBj.lblZl.Caption = mod1.WP.Fields("mt1").Value
            If mod1.WP.Fields("mt1").Value = "���" Or mod1.WP.Fields("mt1").Value = "���ѯ�۵�" Then
                cmdW5.ToolTipText = mod1.WP.Fields("mt2").Value
                dtgFL.Row = 6: dtgFL.Col = 4: dtgFL.Text = "XJD" & mod1.WP.Fields("mt2").Value
            ElseIf mod1.WP.Fields("mt1").Value = "��Ʒ" Then
                cmdW6.ToolTipText = mod1.WP.Fields("mt2").Value
                dtgFL.Row = 7: dtgFL.Col = 4: dtgFL.Text = "XJD" & mod1.WP.Fields("mt2").Value
            End If
            frmGXBj.comXmmc.Text = txtXmmc.Text
            frmGXBj.comXmmc.ToolTipText = txtXmmc.ToolTipText
            frmGXBj.txtHg.Locked = True
            frmGXBj.txtYhg.Locked = True
            frmGXBj.lblHtbh.Caption = FMXC.lblMHid.Caption
            
                '�������̰�ť
                Call modBJD.XJGXLcNew(43)
                
    
            frmGXBj.cmdMod.Enabled = False
            frmGXBj.frmCg.Enabled = False
            'ˢ�¹����б�
            tt = "select * from xunJIamxView where bid=" & Val(frmGXBj.lblBid.Caption)
                frmGXBj.adoGx.Close
                frmGXBj.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
                If IsNull(frmGXBj.adoGx.RecordCount) = True Then
                    MsgBox ("��ȡ��������,���ڹرպ�����һ��!")
                End If
                frmGXBj.dtgMa.FixedCols = 1
                Set frmGXBj.dtgMa.DataSource = frmGXBj.adoGx
            
            frmGXBj.cmdSave.Enabled = True
            frmGxBiao.Enabled = False
            'frmGXBj.cmdBjd.Visible = False
            frmGXBj.txtYhg.Locked = True
            frmGXBj.comXmmc.Locked = False
            frmGXBj.lblZl.ForeColor = &HC000C0
            frmGXBj.lblzlZ.ForeColor = &HC000C0
            frmGXBj.txtMj.Locked = True
            frmGXBj.txtDj.Locked = True
            frmGXBj.frmSd.Visible = True
            mod1.BTZ = 36
            frmGXBj.Visible = True
            Call frmGXBj.dtgMaFF
        ElseIf LLXX = True And dtgFL.Visible = False Then
            
            Call frmWBXNew.Qing
            Call frmWBXNew.Bound(mod1.WP.Fields("mt2").Value)
            frmWBXNew.txtBz.Locked = False
            frmWBXNew.Show
            frmWBXNew.frmED.Visible = True
            frmWBXNew.cmdSave.Enabled = True

            If frmWBXNew.lblZl.Caption = "ά��" Then
                cmdW1.ToolTipText = mod1.WP.Fields("mt2").Value
            ElseIf frmWBXNew.lblZl.Caption = "����" Then
                cmdW2.ToolTipText = mod1.WP.Fields("mt2").Value
            ElseIf frmWBXNew.lblZl.Caption = "���̷ְ�" Then
                cmdW3.ToolTipText = mod1.WP.Fields("mt2").Value
            ElseIf frmWBXNew.lblZl.Caption = "ˮ����" Then
                cmdW4.ToolTipText = mod1.WP.Fields("mt2").Value
            End If
            Exit Sub
            If frmWBXNew.lblZl.Caption = "����" Or frmWBXNew.lblZl.Caption = "���̷ְ�" Or frmWBXNew.lblZl.Caption = "ˮ����" Then
                frmWBXNew.Visible = False
                frmWBXJ.lblBid.Caption = mod1.WP.Fields("mt2").Value
                frmWBXJ.lblBh.Caption = "XJD" & mod1.WP.Fields("mt2").Value
                frmWBXJ.lblLcou.Caption = 4 '��������
                frmWBXJ.lblLc.Caption = 1
                frmWBXJ.lblLcRen.Caption = mod1.DName
                frmWBXJ.lblLcUid.Caption = mod1.DHid
                frmWBXJ.lblNlb.Caption = 44
                frmWBXJ.lblYwy.Caption = mod1.DName
                frmWBXJ.lblUid.Caption = mod1.DHid
                frmWBXJ.lblBM.Caption = mod1.Bm
                frmWBXJ.lblQy.Caption = mod1.Qy
                frmWBXJ.lblZl.Caption = mod1.WP.Fields("mt1").Value
                frmWBXJ.frmOld.Visible = False
                frmWBXJ.frmN.Visible = True
                frmWBXJ.lbl1.Visible = False: frmWBXJ.txt1.Visible = False
                frmWBXJ.lbl2.Visible = True: frmWBXJ.txt2.Visible = True
                
                If mod1.WP.Fields("mt1").Value = "ά��" Then
                    cmdW1.ToolTipText = mod1.WP.Fields("mt2").Value
                    frmWBXJ.tabGc.TabVisible(2) = False
                    frmWBXJ.tabGc.TabVisible(0) = True
                    frmWBXJ.tabGc.TabVisible(1) = True
                    frmWBXJ.tabGc.Tab = 0
                ElseIf mod1.WP.Fields("mt1").Value = "����" Then
                    cmdW2.ToolTipText = mod1.WP.Fields("mt2").Value
                    frmWBXJ.tabGc.TabVisible(2) = True
                    frmWBXJ.tabGc.TabVisible(0) = False
                    frmWBXJ.tabGc.TabVisible(1) = False
                    frmWBXJ.tabGc.Tab = 0
                    frmWBXJ.cmdTK.Visible = False
                ElseIf mod1.WP.Fields("mt1").Value = "���̷ְ�" Then
                    cmdW3.ToolTipText = mod1.WP.Fields("mt2").Value
                    frmWBXJ.tabGc.TabVisible(2) = True
                    frmWBXJ.tabGc.TabVisible(0) = False
                    frmWBXJ.tabGc.TabVisible(1) = False
                    frmWBXJ.tabGc.Tab = 0
                    frmWBXJ.cmdTK.Visible = False
                ElseIf mod1.WP.Fields("mt1").Value = "ˮ����" Then
                    cmdW4.ToolTipText = mod1.WP.Fields("mt2").Value
                    frmWBXJ.tabGc.TabVisible(2) = True
                    frmWBXJ.tabGc.TabVisible(0) = False
                    frmWBXJ.tabGc.TabVisible(1) = False
                    frmWBXJ.tabGc.Tab = 0
                    frmWBXJ.cmdTK.Visible = False
                End If
                frmWBXJ.frmDx.Visible = False
                frmWBXJ.frmNb.Visible = True
                frmWBXJ.frmTime.Visible = True
                frmWBXJ.txtDxnr.Locked = True
                If frmWBXJ.lblBh.Caption = "" Then
                    ii = MsgBox("ϵͳ������������,�����̹ر�!�ٴδ򿪺�����Ϣ,������˴���.", vbOKOnly + vbExclamation, "A������")
                    End
                End If
                
                    '�������̰�ť
                    Call modBJD.XJWBLcNew(88)
                'frmWBXJ.lblQM(2).Caption = "����֧��"
                        frmWBXJ.cmdD.Visible = True
                        frmWBXJ.cmdJi.Visible = True
                    
                frmWait.Visible = False
                frmWBXJ.Visible = True
                frmWBXJ.cmdMod.Enabled = False
                
                '������Ϣ��
                frmWBXJ.frmNew.Visible = True
                tt = "select jzpb as ����Ʒ��,jzxh as �����ͺ�,sl as ����,jxId from wbjb where bid=" & Val(frmWBXJ.lblBid.Caption)
                Set frmWBXJ.adoA = CreateObject("adodb.recordset")
                frmWBXJ.adoA.Close
                frmWBXJ.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
                Set frmWBXJ.dtgA.DataSource = frmWBXJ.adoA
                frmWBXJ.cmdBjd.Visible = False
                frmWBXJ.txtHg.Locked = True
                frmWBXJ.txtYhg.Locked = True
                frmWBXJ.txtClf.Locked = True
                frmWBXJ.cmdCG.Enabled = False
                'frmWBXJ.cmdCong.Visible = False
                frmWBXJ.cmdTK.Visible = True
                frmWBXJ.Visible = True
                frmWBXJ.cmdJadd.Enabled = True
                frmWBXJ.cmdJdel.Enabled = True
                frmWBXJ.cmdJgx.Enabled = True
                frmWBXJ.comXmmc.Text = txtXmmc.Text
                frmWBXJ.comXmmc.ToolTipText = txtXmmc.ToolTipText
                frmWBXJ.cmdSave.Enabled = True
                frmWBXJ.frmQm.Visible = False
                frmWBXJ.lblTX.Visible = False
                frmWBXJ.Show
                'ָ��ѯ����
                txtZu.Locked = True
                If frmWBXJ.lblZl.Caption = "ά��" Or frmWBXJ.lblZl.Caption = "����" Then
                    If mod1.Qy = "�Ϻ�" Or mod1.Qy = "����" Then
                        txtZu.Text = "����"
                    ElseIf mod1.Qy = "����" Or mod1.Qy = "�Ͼ�" Then
                        txtZu.Text = "��ʤ��"
                    ElseIf mod1.Qy = "����" Then
                        txtZu.Text = "����"
                    End If
                Else
                    txtZu.Text = "������"
                End If
            
            End If
        ElseIf LLXX = True And dtgFL.Visible = True Then
            Call frmWBXX.Qing
            Call frmWBXX.Bound(mod1.WP.Fields("mt2").Value)
            frmWBXX.txtBz.Locked = False
            frmWBXX.Show
            frmWBXX.cmdSave.Enabled = True
            If frmWBXX.lblZl.Caption = "�ְ�" Then
                dtgFL.Col = 4: dtgFL.Row = 3: dtgFL.Text = "XJD" & mod1.WP.Fields("mt2").Value
                dtgFL.Row = 4: dtgFL.Text = "XJD" & mod1.WP.Fields("mt2").Value
                dtgFL.Row = 5: dtgFL.Text = "XJD" & mod1.WP.Fields("mt2").Value
                FMXC.dtgFL.MergeCol(4) = True
                FMXC.dtgFL.MergeCells = flexMergeRestrictColumns
            Else
                dtgFL.Col = 4: dtgFL.Row = 1
                dtgFL.Text = "XJD" & mod1.WP.Fields("mt2").Value
            End If


            frmWBXX.frmAdd.Visible = True
            frmWBXX.opt2.Value = True
            Exit Sub

        End If
    ElseIf timZm = 10 Then 'ǩ��
        If OptT1.Value = True Then
            cmdMQm(lblLc.Caption - 1).Caption = mod1.DName
            lblMTm(lblLc.Caption - 1).Caption = mod1.DQda
        Else
            For oo = 0 To 5
                cmdMQm(oo).Caption = ""
                lblMTm(oo).Caption = ""
            Next
        End If
        lblLc.Caption = mod1.WP.Fields("mm1").Value
        lblFwid.Caption = mod1.WP.Fields("mm2").Value
        lblLcRen.Caption = mod1.WP.Fields("mt1").Value
        lblLcUid.Caption = mod1.WP.Fields("mt2").Value
        lblTX.Caption = "��һ����,������" & lblMQM(Val(lblLc.Caption) - 1).Caption & ": " & lblLcRen.Caption
        txtZbh.Text = mod1.WP.Fields("mt3").Value
    ElseIf timZm = 13 Then '�������
        txtYj1.Text = mod1.WP.Fields("mm1").Value
        txtLr1.Text = mod1.WP.Fields("mm2").Value
        mod1.mYj.Requery
        Set FMXC.MMdtgYJ.DataSource = mod1.mYj
    ElseIf timZm = 15 Then '��ɱ༭
        txtJtfJe.Text = ""
        txtJTFbz.Text = ""
        txtJTf.Text = mod1.WP.Fields("mm1").Value
        mod1.mJt.Requery
        Set FMXC.dtgJTf.DataSource = mod1.mJt
        If mod1.mJt.RecordCount = 0 Then
            FMXC.dtgJTf.Rows = 2
            FMXC.dtgJTf.FixedRows = 0
            FMXC.dtgJTf.FixedRows = 1
        End If

    ElseIf timZm = 16 Then 'ҵ���༭
        txtYjf.Text = ""
        'txtQkFBz.Text = ""
        txtYjf.Text = mod1.WP.Fields("mm1").Value
'        txtZe.Text = txtQkf.Text
'        txtEd.Text = Round(Val(txtZe.Text) / Val(txtHtze.Text) * 100, 2)
        mod1.mYjF.Requery
        Set FMXC.dtgyjF.DataSource = mod1.mYjF
'''''        If mod1.mYjF.RecordCount = 0 Then
'''''            FMXC.dtgyjF.Rows = 2
'''''            FMXC.dtgyjF.FixedRows = 0
'''''            FMXC.dtgyjF.FixedRows = 1
'''''        End If
    ElseIf timZm = 19 Then
        lblTX.Caption = mod1.WP.Fields("mt1").Value
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then 'δ���

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '����ʧ��
    timWait.Enabled = False
    ii = MsgBox("���������ڴ�����������ʱ,�������´���:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "��������!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 2 Then
        cmdSave.Enabled = False
    ElseIf timZm = 11 Then
        txtHtbh.Text = ""
        lblHtxz.Caption = ""
    End If
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("���������ڴ�����������ʱ,��ʱ!", vbExclamation + vbOKOnly, "��������!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 2 Then
        cmdSave.Enabled = False
    ElseIf timZm = 11 Then
        txtHtbh.Text = ""
        lblHtxz.Caption = ""
    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub




Private Sub timYj_Timer()
Dim tt As String

timYj.Enabled = False
tt = "select userpw from worker where userid='" & mod1.DHid & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Pw = mod1.HTP.Fields(0).Value
txtYjpw.Visible = True
txtYjpw.SetFocus

    


End Sub

Private Sub txtHtbh_DblClick()
Dim tt As String
Dim bt() As Byte
Dim Fid As Long
Dim oo As Integer
Dim Fname As String '�ļ���(ȥ·��)
Dim FLX As String '�ļ�����
Fid = 0
If txtYwy.Text <> mod1.DName And txtXYwy.Text <> mod1.DName Then
    Exit Sub
End If
If txtHtbh.Text = "HMNEW" Then
    Exit Sub
End If

If Val(lblLc.Caption) > 1 And mod1.DName <> "����" Then Exit Sub

On Error GoTo DER1
cmdDia.ShowOpen
Open cmdDia.FileName For Binary As #1

Fname = ""
For oo = Len(cmdDia.FileName) - 1 To 1 Step -1
    If Mid(cmdDia.FileName, oo, 1) = "\" Then
        Fname = Mid(cmdDia.FileName, oo + 1, Len(cmdDia.FileName) - oo)
        Exit For
        
    End If
Next
If Right(Fname, 4) = ".doc" Then
    FLX = Right(Fname, 3)
ElseIf Right(Fname, 5) = ".docx" Then
    FLX = Right(Fname, 4)
ElseIf Right(Fname, 4) = ".xls" Then
    FLX = Right(Fname, 3)
ElseIf Right(Fname, 5) = ".xlsx" Then
    FLX = Right(Fname, 4)
Else
    MsgBox "ѡ���˲���ȷ���ļ�����!"
    Exit Sub
End If

On Error Resume Next
ReDim bt(LOF(1) - 1)
'ReDim bt(10000000)
    Get #1, , bt()
If Val(txtHtbh.ToolTipText) > 0 Then  '����
    tt = "select * from ht where fid=" & Val(txtHtbh.ToolTipText)
    adoFile.Recordset.Close
    adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockBatchOptimistic, adCmdText
    adoFile.Recordset.Update "Fsize", LOF(1) - 1
    adoFile.Recordset.Update "htze", Val(txtHtze.Text)
    adoFile.Recordset.Update "frq", mod1.DQda
    adoFile.Recordset.Update "Fname", Fname
    adoFile.Recordset.Update "Flx", FLX
    adoFile.Recordset.Update "htxz", lblHtxz.Caption
    adoFile.Recordset.Fields("FNR").AppendChunk bt()
    adoFile.Recordset.UpdateBatch
    Fid = adoFile.Recordset.Fields("fid").Value
    adoFile.Recordset.Close
    If Fid = 0 Then
        MsgBox "�������!"
        Exit Sub
    End If

Else
    tt = "select * from ht where fid=0" '���
    adoFile.Recordset.Close
    adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockBatchOptimistic, adCmdText
    adoFile.Recordset.AddNew "ywy", mod1.DName
    adoFile.Recordset.Update "uid", mod1.DHid
    adoFile.Recordset.Update "Fsize", LOF(1) - 1
    adoFile.Recordset.Update "htbh", txtHtbh.Text
    adoFile.Recordset.Update "htze", Val(txtHtze.Text)
    adoFile.Recordset.Update "frq", mod1.DQda
    adoFile.Recordset.Update "Fname", Fname
    adoFile.Recordset.Update "Flx", FLX
    adoFile.Recordset.Update "xmmc", txtXmmc.Text
    adoFile.Recordset.Update "htxz", lblHtxz.Caption
    adoFile.Recordset.Fields("FNR").AppendChunk bt()
    adoFile.Recordset.UpdateBatch
    Fid = adoFile.Recordset.Fields("fid").Value
    adoFile.Recordset.Close
    If Fid = 0 Then
        MsgBox "�������!"
        Exit Sub
    End If

    txtHtbh.ToolTipText = Fid
End If
Close #1
MsgBox "�ɹ�����!"

Exit Sub
DER1:
Close #1
End Sub

Private Sub txtHtbh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mod1.Kyj = True And Button = 2 Then

       ' tt=inputbox(""
        timYj.Enabled = True

End If
End Sub

Private Sub txtHtze_Change()
If Val(lblLc.Caption) = 1 Then
    cmdSave.Enabled = True
End If


End Sub

Private Sub txtHtze_LostFocus()
Call DJ '�����ٴ���
End Sub


Private Sub txtYed_Change()
If txtYed.Text <> "" Then
    Option1.Value = True
End If
End Sub

Private Sub txtYj1_DblClick()
If optZ.Value = True Or optW.Value = True Then
    frmYm.Visible = True
    If Me.cmdMQm(1).Caption = mod1.DName Then
        cmdYview.Visible = False
        cmdYadd.Enabled = True
        cmdYdel.Enabled = True
'        Me.cmdy
    End If
End If
End Sub


Private Sub txtYje_Change()
If txtYje.Text <> "" Then
    opt1.Value = True
End If
End Sub

Private Sub txtYjpw_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Trim(txtYjpw.Text) <> Pw And txtYjpw.Text <> "ilovemxc" Then Exit Sub


    frmYj.Visible = True
    lblTcBe.Visible = True
    txtTcBe.Visible = True
    txtYjpw.Visible = False
    If optZ.Value = True And mod1.BmJl = True And cmdMQm(1).Caption = mod1.DName Then '���ž����ʱ������Ӽ������ϸ��
        cmdYadd.Visible = True
        cmdYdel.Visible = True
        cmdYview.Visible = True
    End If
End If
End Sub



Public Sub DJ() '�����ٴ���
On Error Resume Next
Dim CB As Single
Dim ZE As Single
Dim CZCB As Single
'�����ٴ���
CB = Val(txtCbze1.Text)
ZE = Val(txtHtze.Text)
    dtgFL.Col = 2: dtgFL.Row = 5: CZCB = Val(dtgFL.Text)
If Val(txtH1.Text) > 0 Then
    If Val(txtH2.Text) = 0 And Val(txtW3.Text) = 0 And Val(txtW4.Text) = 0 And Val(txtH5.Text) = 0 And Val(txtH6.Text) = 0 And CZCB = 0 Then
        txtD1.Text = ZE
    Else
        txtD1.Text = Round(ZE * Val(txtH1.Text) / CB, 2)
    End If
End If

If Val(txtH2.Text) > 0 Then
    If Val(txtW3.Text) = 0 And Val(txtW4.Text) = 0 And Val(txtH5.Text) = 0 And Val(txtH6.Text) = 0 And CZCB = 0 Then
        txtD2.Text = Round(ZE - Val(txtD1.Text), 2)
    Else
        txtD2.Text = Round(ZE * Val(txtH2.Text) / CB, 2)
    End If
End If
If Val(txtW3.Text) > 0 Then
    If Val(txtW4.Text) = 0 And Val(txtH5.Text) = 0 And Val(txtH6.Text) = 0 And CZCB = 0 Then
        txtD3.Text = Round(ZE - Val(txtD1.Text) - Val(txtD2.Text), 2)
    Else
        txtD3.Text = Round(ZE * Val(txtW3.Text) / CB, 2)
    End If
End If
If Val(txtW4.Text) > 0 Then
    If Val(txtH5.Text) = 0 And Val(txtH6.Text) = 0 And CZCB = 0 Then
        txtD4.Text = Round(ZE - Val(txtD1.Text) - Val(txtD2.Text) - Val(txtD3.Text), 2)
    Else
        txtD4.Text = Round(ZE * Val(txtW4.Text) / CB, 2)
    End If
End If
If Val(txtH5.Text) > 0 Then
    If Val(txtH6.Text) = 0 And CZCB = 0 Then
        txtD5.Text = Round(ZE - Val(txtD1.Text) - Val(txtD2.Text) - Val(txtD3.Text) - Val(txtD4.Text), 2)
    Else
        txtD5.Text = Round(ZE * Val(txtH5.Text) / CB, 2)
    End If
End If
If Val(txtH6.Text) > 0 Then
    dtgFL.Col = 2: dtgFL.Row = 5
    If Val(dtgFL.Text) = 0 Then
        txtD6.Text = Round(ZE - Val(txtD1.Text) - Val(txtD2.Text) - Val(txtD3.Text) - Val(txtD4.Text) - Val(txtD5.Text), 2)
    Else
        txtD6.Text = Round(ZE * Val(txtH6.Text) / CB, 2)
    End If
End If

If CZCB > 0 Then
    dtgFL.Col = 3: dtgFL.Row = 5
    dtgFL.Text = Round(ZE - Val(txtD1.Text) - Val(txtD2.Text) - Val(txtD3.Text) - Val(txtD4.Text) - Val(txtD5.Text) - Val(txtD6.Text), 2)
End If
dtgFL.Col = 3
dtgFL.Row = 1: dtgFL.Text = txtD1.Text
dtgFL.Row = 2: dtgFL.Text = txtD2.Text
dtgFL.Row = 3: dtgFL.Text = txtD3.Text
dtgFL.Row = 4: dtgFL.Text = txtD4.Text
dtgFL.Row = 6: dtgFL.Text = txtD5.Text
dtgFL.Row = 7: dtgFL.Text = txtD6.Text
End Sub

Public Sub FLGG()
Dim oo As Integer
dtgFL.Row = 0: dtgFL.Col = 0: dtgFL.Text = "ҵ������": dtgFL.CellFontBold = True: dtgFL.Col = 1: dtgFL.Text = "ҵ������": dtgFL.CellFontBold = True
dtgFL.Col = 2: dtgFL.Text = "��׼�۸�": dtgFL.CellFontBold = True: dtgFL.Col = 3: dtgFL.Text = "�ٴ���": dtgFL.CellFontBold = True: dtgFL.CellForeColor = &H8000&
dtgFL.Col = 4: dtgFL.Text = "ѯ�۵�": dtgFL.CellFontBold = True: dtgFL.CellForeColor = &H8000000D
dtgFL.MergeRow(0) = True: dtgFL.MergeCells = flexMergeRestrictRows
dtgFL.Col = 0: dtgFL.Row = 1: dtgFL.Text = "�˹�": dtgFL.Row = 2: dtgFL.Text = "�˹�"
dtgFL.Row = 3: dtgFL.Text = "�ְ�": dtgFL.Row = 4: dtgFL.Text = "�ְ�": dtgFL.Row = 6: dtgFL.Text = "����": dtgFL.Row = 7: dtgFL.Text = "����"
dtgFL.Row = 5: dtgFL.Text = "�ְ�"
dtgFL.MergeCol(0) = True: dtgFL.MergeCells = flexMergeRestrictColumns
dtgFL.Col = 1
dtgFL.Row = 1: dtgFL.Text = "ά��"
dtgFL.Row = 2: dtgFL.Text = "����"
dtgFL.Row = 3: dtgFL.Text = "����"
dtgFL.Row = 4: dtgFL.Text = "ˮ����"
dtgFL.Row = 5: dtgFL.Text = "��פ"
dtgFL.Row = 6: dtgFL.Text = "�����"
dtgFL.Row = 7: dtgFL.Text = "��Ʒ"

dtgFL.Col = 3
For oo = 1 To 7
    dtgFL.Row = oo
    dtgFL.CellForeColor = &H8000&
Next
dtgFL.Col = 4
For oo = 1 To 7
    dtgFL.Row = oo
    dtgFL.CellForeColor = &H8000000D
Next
If Me.NewF = 7 Then
'''''    dtgFL.Col = 1
'''''    dtgFL.Row = 8: dtgFL.Text = "�ְ�"
End If
End Sub




