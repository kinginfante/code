VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMNew 
   Caption         =   "�����°��ͬ����"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Frame frmQM 
      BackColor       =   &H00C0FFC0&
      Caption         =   "������"
      ForeColor       =   &H000000FF&
      Height          =   1725
      Left            =   2730
      TabIndex        =   231
      Top             =   7440
      Width           =   6045
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "����"
         Height          =   285
         Left            =   5130
         Style           =   1  'Graphical
         TabIndex        =   235
         Top             =   1290
         Width           =   735
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�ܾ�"
         Height          =   195
         Left            =   5130
         TabIndex        =   234
         Top             =   900
         Width           =   675
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ͬ��"
         Height          =   225
         Left            =   5130
         TabIndex        =   233
         Top             =   510
         Width           =   705
      End
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1245
         Left            =   90
         TabIndex        =   232
         Top             =   330
         Width           =   4965
      End
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   5
      Left            =   5910
      TabIndex        =   167
      Top             =   8370
      Width           =   945
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   4
      Left            =   4890
      TabIndex        =   164
      Top             =   8370
      Width           =   945
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   3
      Left            =   3900
      TabIndex        =   161
      Top             =   8370
      Width           =   945
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   2
      Left            =   2910
      TabIndex        =   158
      Top             =   8370
      Width           =   945
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   1
      Left            =   1890
      TabIndex        =   155
      Top             =   8370
      Width           =   945
   End
   Begin VB.Timer timWait 
      Interval        =   1000
      Left            =   7860
      Top             =   7860
   End
   Begin VB.Timer timQuit 
      Interval        =   1000
      Left            =   8670
      Top             =   7860
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "�ύ"
      Height          =   585
      Left            =   13320
      Picture         =   "frmHTNew.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   8580
      Width           =   675
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   585
      Left            =   14640
      Picture         =   "frmHTNew.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   8580
      Width           =   585
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "�޸�"
      Height          =   585
      Left            =   12630
      Picture         =   "frmHTNew.frx":076C
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   8580
      Width           =   645
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ��"
      Enabled         =   0   'False
      Height          =   585
      Left            =   13980
      Picture         =   "frmHTNew.frx":0BAE
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   8580
      Width           =   645
   End
   Begin VB.Frame frmZt 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   1005
      Left            =   10650
      TabIndex        =   116
      Top             =   8160
      Visible         =   0   'False
      Width           =   1185
      Begin VB.OptionButton optW 
         Caption         =   "ִ�����"
         Height          =   225
         Left            =   60
         TabIndex        =   120
         Top             =   510
         Width           =   1035
      End
      Begin VB.OptionButton optZ 
         Caption         =   "ִ�н׶�"
         Height          =   225
         Left            =   60
         TabIndex        =   119
         Top             =   300
         Width           =   1035
      End
      Begin VB.OptionButton optP 
         Caption         =   "����׶�"
         Height          =   180
         Left            =   60
         TabIndex        =   118
         Top             =   780
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optG 
         Caption         =   "�� �� ��"
         Height          =   195
         Left            =   60
         TabIndex        =   117
         Top             =   90
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   0
      Left            =   870
      TabIndex        =   111
      Top             =   8370
      Width           =   945
   End
   Begin VB.CommandButton cmdPje 
      Caption         =   "������"
      Height          =   1095
      Left            =   450
      TabIndex        =   110
      Top             =   8070
      Width           =   345
   End
   Begin VB.CommandButton cmdCong 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��������"
      Height          =   1095
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   8070
      Width           =   345
   End
   Begin TabDlg.SSTab tabHt 
      Height          =   7905
      Left            =   -60
      TabIndex        =   0
      Top             =   -120
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   13944
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "����"
      TabPicture(0)   =   "frmHTNew.frx":0D38
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label49"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblHtxz"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label29"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label38"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label44"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label25"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label15"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label5"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label18"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label19"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label20"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label9"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label17"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label24"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label26"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblJlr"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label6"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label7"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label30"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Line2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Line3"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "MMdtgFk"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtBz"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "frmFX"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "frmFk"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "frmYj"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtJlr2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtQt2"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtCbze2"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtYf2"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtFbje2"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtHtrq"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtZe"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtEd"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "comQy"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtXYwy"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtHtbh"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cmdWb"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtHtze"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtRgf1"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtCLF1"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtFbje1"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtYf1"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtQt1"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtClcb1"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Frame1"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtCbze1"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtADR"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtJlr1"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtKhdm"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtXMMC"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txtKhmc"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txtTcRQ"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtYwy"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "Frame3"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "frmHide"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "frmYM"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txtClcb2"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txtRGF2"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "cmdHt"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txtMxmmc"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).ControlCount=   68
      TabCaption(1)   =   "��������"
      TabPicture(1)   =   "frmHTNew.frx":0D54
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tabGc"
      Tab(1).Control(1)=   "Command1"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtMxmmc 
         Height          =   315
         Left            =   570
         TabIndex        =   238
         Text            =   "Text1"
         Top             =   150
         Width           =   2865
      End
      Begin VB.CommandButton cmdHt 
         BackColor       =   &H008080FF&
         Caption         =   "BH"
         Height          =   225
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   236
         Top             =   1680
         Width           =   405
      End
      Begin VB.TextBox txtRGF2 
         Height          =   285
         Left            =   13440
         Locked          =   -1  'True
         TabIndex        =   154
         Top             =   570
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtClcb2 
         Height          =   315
         Left            =   13050
         Locked          =   -1  'True
         TabIndex        =   153
         Top             =   1515
         Width           =   1185
      End
      Begin VB.Frame frmYM 
         BackColor       =   &H8000000D&
         Caption         =   "����Ԥ��֧�����"
         Height          =   2055
         Left            =   4560
         TabIndex        =   2
         Top             =   4950
         Visible         =   0   'False
         Width           =   4665
         Begin VB.CommandButton cmdYdel 
            Caption         =   "ɾ��"
            Height          =   285
            Left            =   3960
            TabIndex        =   7
            Top             =   1170
            Width           =   585
         End
         Begin VB.CommandButton cmdYadd 
            Caption         =   "���"
            Height          =   315
            Left            =   3960
            TabIndex        =   6
            Top             =   810
            Width           =   585
         End
         Begin VB.TextBox txtYingFu 
            Height          =   270
            Left            =   2850
            TabIndex        =   5
            Top             =   1620
            Width           =   1035
         End
         Begin VB.TextBox txtFED 
            Height          =   285
            Left            =   930
            TabIndex        =   4
            Top             =   1620
            Width           =   645
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "�ر�"
            Height          =   285
            Left            =   3960
            TabIndex        =   3
            Top             =   1590
            Width           =   615
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgYJ 
            Height          =   1275
            Left            =   30
            TabIndex        =   8
            Top             =   210
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   2249
            _Version        =   393216
            BackColorBkg    =   -2147483635
            SelectionMode   =   1
            BorderStyle     =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label39 
            BackColor       =   &H8000000D&
            Caption         =   "֧�����"
            Height          =   225
            Left            =   1980
            TabIndex        =   12
            Top             =   1650
            Width           =   915
         End
         Begin VB.Label Label40 
            BackColor       =   &H8000000D&
            Caption         =   "%"
            Height          =   255
            Left            =   1680
            TabIndex        =   11
            Top             =   1650
            Width           =   195
         End
         Begin VB.Label Label41 
            BackColor       =   &H8000000D&
            Caption         =   "�տ���"
            Height          =   255
            Left            =   90
            TabIndex        =   10
            Top             =   1650
            Width           =   825
         End
         Begin VB.Label lblyjFF 
            Caption         =   "lblYjff"
            Height          =   255
            Left            =   3600
            TabIndex        =   9
            Top             =   330
            Visible         =   0   'False
            Width           =   1125
         End
      End
      Begin VB.Frame frmHide 
         Caption         =   "frmHid"
         Height          =   2775
         Left            =   4920
         TabIndex        =   41
         Top             =   330
         Visible         =   0   'False
         Width           =   4935
         Begin VB.Label lblBm 
            Caption         =   "lblBm"
            Height          =   225
            Left            =   150
            TabIndex        =   52
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblQy 
            Caption         =   "lblQy"
            Height          =   255
            Left            =   2940
            TabIndex        =   51
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label lblLc 
            Caption         =   "lblLc"
            Height          =   315
            Left            =   150
            TabIndex        =   50
            Top             =   600
            Width           =   645
         End
         Begin VB.Label lblNlb 
            Caption         =   "lblNlb"
            Height          =   225
            Left            =   1470
            TabIndex        =   49
            Top             =   570
            Width           =   645
         End
         Begin VB.Label lblLcRen 
            Caption         =   "lblLcRen"
            Height          =   285
            Left            =   150
            TabIndex        =   48
            Top             =   810
            Width           =   795
         End
         Begin VB.Label lblLcUid 
            Caption         =   "lblLcUid"
            Height          =   285
            Left            =   180
            TabIndex        =   47
            Top             =   1020
            Width           =   885
         End
         Begin VB.Label lblFwid 
            Caption         =   "lblFwid"
            Height          =   255
            Left            =   1380
            TabIndex        =   46
            Top             =   210
            Width           =   885
         End
         Begin VB.Label lblUid 
            Caption         =   "lblUid"
            Height          =   255
            Left            =   2580
            TabIndex        =   45
            Top             =   780
            Width           =   975
         End
         Begin VB.Label lblYwy 
            Caption         =   "lblYwy"
            Height          =   285
            Left            =   2520
            TabIndex        =   44
            Top             =   450
            Width           =   765
         End
         Begin VB.Label lblLcou 
            Caption         =   "lblLcou"
            Height          =   255
            Left            =   1500
            TabIndex        =   43
            Top             =   960
            Width           =   1185
         End
         Begin VB.Label lblPwf 
            Caption         =   "lblPwf"
            Height          =   225
            Left            =   2520
            TabIndex        =   42
            Top             =   1080
            Width           =   1185
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "�ͻ�������:"
         Height          =   3705
         Left            =   5070
         TabIndex        =   126
         Top             =   3630
         Width           =   5265
         Begin VB.CommandButton cmdW3 
            Caption         =   "ѯ�۵�"
            Height          =   285
            Left            =   4320
            TabIndex        =   148
            Top             =   1356
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton cmdW6 
            Caption         =   "ѯ�۵�"
            Height          =   255
            Left            =   4320
            TabIndex        =   147
            Top             =   2670
            Width           =   765
         End
         Begin VB.CommandButton cmdW5 
            Caption         =   "ѯ�۵�"
            Height          =   285
            Left            =   4320
            TabIndex        =   146
            Top             =   2232
            Width           =   765
         End
         Begin VB.CommandButton cmdW2 
            Caption         =   "ѯ�۵�"
            Height          =   285
            Left            =   4320
            TabIndex        =   145
            Top             =   948
            Width           =   765
         End
         Begin VB.CommandButton cmdW1 
            Caption         =   "ѯ�۵�"
            Height          =   285
            Left            =   4320
            TabIndex        =   144
            Top             =   510
            Width           =   765
         End
         Begin VB.TextBox txtH6 
            Height          =   270
            Left            =   3090
            TabIndex        =   143
            Top             =   2670
            Width           =   915
         End
         Begin VB.TextBox txtW6 
            Height          =   270
            Left            =   1890
            TabIndex        =   142
            Top             =   2670
            Width           =   915
         End
         Begin VB.TextBox txtH5 
            Height          =   270
            Left            =   3090
            TabIndex        =   141
            Top             =   2238
            Width           =   915
         End
         Begin VB.TextBox txtW5 
            Height          =   270
            Left            =   1890
            TabIndex        =   140
            Top             =   2238
            Width           =   915
         End
         Begin VB.TextBox txtW4 
            Height          =   270
            Left            =   1890
            TabIndex        =   139
            Top             =   1806
            Width           =   915
         End
         Begin VB.TextBox txtH3 
            Height          =   270
            Left            =   3090
            TabIndex        =   138
            Top             =   1374
            Width           =   915
         End
         Begin VB.TextBox txtW3 
            Height          =   270
            Left            =   1860
            TabIndex        =   137
            Top             =   1374
            Width           =   915
         End
         Begin VB.TextBox txtH2 
            Height          =   270
            Left            =   1860
            TabIndex        =   136
            Top             =   945
            Width           =   2175
         End
         Begin VB.TextBox txtH1 
            Height          =   270
            Left            =   1860
            TabIndex        =   135
            Top             =   510
            Width           =   2175
         End
         Begin VB.CheckBox chkF 
            Caption         =   "���Ϸ�(��Ʒ)"
            Height          =   225
            Left            =   120
            TabIndex        =   132
            Top             =   2700
            Width           =   1425
         End
         Begin VB.CheckBox chkE 
            Caption         =   "���Ϸ�(���)"
            Height          =   285
            Left            =   120
            TabIndex        =   131
            Top             =   2232
            Width           =   1575
         End
         Begin VB.CheckBox chkD 
            Caption         =   "�˹���(ˮ����)"
            Height          =   225
            Left            =   120
            TabIndex        =   130
            ToolTipText     =   "�빤�̲��޹ص��˹�(��ˮ����)"
            Top             =   1824
            Width           =   1605
         End
         Begin VB.CheckBox chkC 
            Caption         =   "�˹���(���̷ְ�)"
            Height          =   195
            Left            =   120
            TabIndex        =   129
            ToolTipText     =   "�ɹ��̶�������"
            Top             =   1446
            Width           =   1785
         End
         Begin VB.CheckBox chkB 
            Caption         =   "�˹���(����)"
            Height          =   315
            Left            =   120
            TabIndex        =   128
            ToolTipText     =   "�ɹ���һ������,���д��޻�һ����ά��,�����ڲ�����9����"
            Top             =   948
            Width           =   1395
         End
         Begin VB.Label Label32 
            Caption         =   "�˼۳ɱ�"
            Height          =   225
            Left            =   3180
            TabIndex        =   134
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label31 
            Caption         =   "Ԥ���ɱ�:"
            Height          =   225
            Left            =   1830
            TabIndex        =   133
            Top             =   180
            Width           =   1035
         End
         Begin MSForms.CheckBox chkA 
            Height          =   255
            Left            =   120
            TabIndex        =   127
            ToolTipText     =   "�ɹ���һ������,�����ڳ���9����"
            Top             =   510
            Width           =   1635
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2884;450"
            Value           =   "0"
            Caption         =   "�˹���(ά��)"
            FontName        =   "����"
            FontHeight      =   180
            FontCharSet     =   134
            FontPitchAndFamily=   34
         End
      End
      Begin VB.TextBox txtYwy 
         Height          =   270
         Left            =   8610
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   1080
         Width           =   1305
      End
      Begin VB.TextBox txtTcRQ 
         Height          =   315
         Left            =   10650
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "���ȡ������"
         Top             =   6960
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.ComboBox txtKhmc 
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   81
         ToolTipText     =   "�����б���ѡ��ͻ�"
         Top             =   630
         Width           =   3345
      End
      Begin VB.TextBox txtXMMC 
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   600
         Width           =   3555
      End
      Begin VB.TextBox txtKhdm 
         Height          =   270
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   1140
         Width           =   3315
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
         Left            =   11730
         TabIndex        =   78
         Top             =   3782
         Width           =   1245
      End
      Begin VB.TextBox txtADR 
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   2160
         Width           =   3555
      End
      Begin VB.TextBox txtCbze1 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   11730
         Locked          =   -1  'True
         TabIndex        =   76
         ToolTipText     =   "Ԥ��"
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Frame Frame1 
         Caption         =   "��Ʊ���ͣ�"
         Height          =   765
         Left            =   240
         TabIndex        =   72
         Top             =   6570
         Width           =   4035
         Begin VB.OptionButton optLc 
            Caption         =   "����Ʊ"
            Height          =   195
            Left            =   2370
            TabIndex        =   75
            Top             =   300
            Width           =   1065
         End
         Begin VB.OptionButton optLb 
            Caption         =   "��ҵ��Ʊ"
            Height          =   195
            Left            =   1260
            TabIndex        =   74
            Top             =   300
            Width           =   1065
         End
         Begin VB.OptionButton optLa 
            Caption         =   "��ֵ��Ʊ"
            Height          =   195
            Left            =   180
            TabIndex        =   73
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.TextBox txtClcb1 
         Height          =   285
         Left            =   11730
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   1530
         Width           =   1215
      End
      Begin VB.TextBox txtQt1 
         Height          =   285
         Left            =   11730
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   3330
         Width           =   2535
      End
      Begin VB.TextBox txtYf1 
         Height          =   285
         Left            =   11730
         TabIndex        =   69
         ToolTipText     =   "Ԥ��"
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtFbje1 
         Height          =   285
         Left            =   11730
         Locked          =   -1  'True
         TabIndex        =   68
         ToolTipText     =   "Ԥ��"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtCLF1 
         Height          =   285
         Left            =   11730
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   2445
         Width           =   2505
      End
      Begin VB.TextBox txtRgf1 
         Height          =   315
         Left            =   11730
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1980
         Width           =   2475
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
         Left            =   1440
         TabIndex        =   65
         ToolTipText     =   "���ڸ�����ϸ��ȷ����ͬ�ܽ��"
         Top             =   3090
         Width           =   3345
      End
      Begin VB.CommandButton cmdWb 
         Caption         =   "��Ŀ����"
         Height          =   315
         Left            =   1470
         TabIndex        =   64
         Top             =   2430
         Width           =   3375
      End
      Begin VB.TextBox txtHtbh 
         Height          =   270
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   1650
         Width           =   3315
      End
      Begin VB.TextBox txtXYwy 
         Height          =   315
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   1095
         Width           =   1245
      End
      Begin VB.ComboBox comQy 
         Height          =   300
         ItemData        =   "frmHTNew.frx":0D70
         Left            =   8970
         List            =   "frmHTNew.frx":0D72
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   1575
         Width           =   945
      End
      Begin VB.TextBox txtEd 
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   9030
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   3090
         Width           =   885
      End
      Begin VB.TextBox txtZe 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   3090
         Width           =   1515
      End
      Begin VB.TextBox txtHtrq 
         Height          =   315
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   1590
         Width           =   1815
      End
      Begin VB.TextBox txtFbje2 
         Height          =   315
         Left            =   13020
         Locked          =   -1  'True
         TabIndex        =   57
         ToolTipText     =   "ʵ��"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtYf2 
         Height          =   315
         Left            =   13020
         Locked          =   -1  'True
         TabIndex        =   56
         ToolTipText     =   "ʵ��"
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtCbze2 
         Height          =   315
         Left            =   13050
         Locked          =   -1  'True
         TabIndex        =   55
         ToolTipText     =   "ʵ��"
         Top             =   1080
         Width           =   1185
      End
      Begin VB.TextBox txtQt2 
         Height          =   285
         Left            =   13260
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   4230
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtJlr2 
         Height          =   285
         Left            =   13020
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3780
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��ӡ"
         Height          =   585
         Left            =   -60420
         Picture         =   "frmHTNew.frx":0D74
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   7260
         Width           =   645
      End
      Begin VB.Frame frmYj 
         Height          =   2385
         Left            =   10590
         TabIndex        =   27
         Top             =   4560
         Width           =   4095
         Begin VB.TextBox txtLr2 
            Height          =   285
            Left            =   2460
            Locked          =   -1  'True
            TabIndex        =   34
            ToolTipText     =   "ʵ��"
            Top             =   630
            Width           =   1215
         End
         Begin VB.TextBox txtYj2 
            Height          =   285
            Left            =   2460
            Locked          =   -1  'True
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtYj1 
            Height          =   285
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   1185
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
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   31
            ToolTipText     =   "Ԥ��"
            Top             =   630
            Width           =   1185
         End
         Begin VB.TextBox txtTc2 
            Height          =   285
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   2010
            Width           =   1305
         End
         Begin VB.TextBox txtTcBe 
            Height          =   285
            Left            =   990
            TabIndex        =   29
            Text            =   "6"
            Top             =   1650
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CommandButton cmdCount 
            Caption         =   "����"
            Height          =   315
            Left            =   1590
            TabIndex        =   28
            Top             =   1650
            Visible         =   0   'False
            Width           =   705
         End
         Begin MSComCtl2.UpDown UpDa 
            Height          =   315
            Left            =   1320
            TabIndex        =   35
            Top             =   1650
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   503
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label lblYj 
            Caption         =   "��    ��"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   300
            Width           =   975
         End
         Begin VB.Label lblTC 
            Caption         =   "��    ��"
            Height          =   195
            Left            =   60
            TabIndex        =   38
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label lblLr 
            Caption         =   "�� �� 2"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   690
            Width           =   915
         End
         Begin VB.Label lblTcBe 
            Caption         =   "��ɱ���"
            Height          =   195
            Left            =   60
            TabIndex        =   36
            Top             =   1710
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Frame frmFk 
         Height          =   555
         Left            =   240
         TabIndex        =   18
         Top             =   5670
         Width           =   4245
         Begin VB.TextBox txtYed 
            Height          =   270
            Left            =   3150
            TabIndex        =   20
            Top             =   150
            Width           =   795
         End
         Begin VB.TextBox txtYrq 
            Height          =   300
            Left            =   900
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   150
            Width           =   1005
         End
         Begin MSComCtl2.DTPicker dtpYf 
            Height          =   315
            Left            =   900
            TabIndex        =   21
            Top             =   150
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   16711680
            CalendarTrailingForeColor=   8454016
            Format          =   149094401
            CurrentDate     =   38797
         End
         Begin VB.Label Label33 
            Caption         =   "Ӧ������"
            Height          =   285
            Left            =   60
            TabIndex        =   25
            Top             =   180
            Width           =   735
         End
         Begin VB.Label Label34 
            Caption         =   "�տ���"
            Height          =   255
            Left            =   2310
            TabIndex        =   24
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label37 
            Caption         =   "%"
            Height          =   255
            Left            =   4050
            TabIndex        =   23
            Top             =   180
            Width           =   435
         End
         Begin VB.Label lblFid 
            Caption         =   "lblFid"
            Height          =   165
            Left            =   3600
            TabIndex        =   22
            Top             =   360
            Visible         =   0   'False
            Width           =   945
         End
      End
      Begin VB.Frame frmFX 
         Height          =   1605
         Left            =   4320
         TabIndex        =   13
         Top             =   3720
         Width           =   585
         Begin VB.CommandButton cmdGx 
            Caption         =   "����"
            Height          =   315
            Left            =   0
            TabIndex        =   17
            Top             =   1230
            Width           =   525
         End
         Begin VB.CommandButton cmdQing 
            Caption         =   "���"
            Height          =   345
            Left            =   0
            TabIndex        =   16
            Top             =   120
            Width           =   525
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "���"
            Height          =   375
            Left            =   0
            TabIndex        =   15
            Top             =   450
            Width           =   525
         End
         Begin VB.CommandButton cmdDe 
            Caption         =   "ɾ��"
            Height          =   375
            Left            =   0
            TabIndex        =   14
            Top             =   840
            Width           =   525
         End
      End
      Begin VB.TextBox txtBz 
         Height          =   465
         Left            =   6360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   2580
         Width           =   3525
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgFk 
         Height          =   1875
         Left            =   150
         TabIndex        =   26
         Top             =   3690
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   3307
         _Version        =   393216
         FillStyle       =   1
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin TabDlg.SSTab tabGc 
         Height          =   7605
         Left            =   -74970
         TabIndex        =   149
         Top             =   -60
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   13414
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "�걣"
         TabPicture(0)   =   "frmHTNew.frx":13DE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "frmgc(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "����"
         TabPicture(1)   =   "frmHTNew.frx":13FA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frmgc(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "���"
         TabPicture(2)   =   "frmHTNew.frx":1416
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "frmgc(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "��Ʒ"
         TabPicture(3)   =   "frmHTNew.frx":1432
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "frmgc(3)"
         Tab(3).Control(1)=   "VScroll1"
         Tab(3).ControlCount=   2
         TabCaption(4)   =   "���̷ְ�"
         TabPicture(4)   =   "frmHTNew.frx":144E
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "frmgc(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "ˮ����"
         TabPicture(5)   =   "frmHTNew.frx":146A
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "frmgc(5)"
         Tab(5).ControlCount=   1
         Begin VB.Frame frmgc 
            Caption         =   "Frame4"
            Height          =   7275
            Index           =   5
            Left            =   -74970
            TabIndex        =   229
            Top             =   30
            Width           =   15195
            Begin VB.TextBox txtWBNR 
               Height          =   7245
               Left            =   0
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   230
               Top             =   -30
               Width           =   15195
            End
         End
         Begin VB.Frame frmgc 
            Caption         =   "Frame4"
            Height          =   7275
            Index           =   4
            Left            =   -74970
            TabIndex        =   227
            Top             =   30
            Width           =   15195
            Begin VB.TextBox txtFBNR 
               Height          =   7245
               Left            =   0
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   228
               Top             =   0
               Width           =   15165
            End
         End
         Begin VB.Frame frmgc 
            Caption         =   "Frame4"
            Height          =   7275
            Index           =   3
            Left            =   -75000
            TabIndex        =   218
            Top             =   30
            Width           =   15225
            Begin VB.TextBox txtCL 
               Height          =   315
               Left            =   9480
               TabIndex        =   222
               Top             =   5970
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.CommandButton Command2 
               Caption         =   "ɾ��"
               Height          =   315
               Left            =   14250
               TabIndex        =   221
               Top             =   5970
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.TextBox txtCj 
               Height          =   345
               Left            =   11880
               TabIndex        =   220
               Top             =   5970
               Width           =   1455
            End
            Begin VB.CommandButton cmdCGX 
               Caption         =   "����"
               Height          =   315
               Left            =   13500
               TabIndex        =   219
               Top             =   5970
               Width           =   675
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgCP 
               Height          =   5865
               Left            =   0
               TabIndex        =   223
               Top             =   0
               Width           =   15225
               _ExtentX        =   26855
               _ExtentY        =   10345
               _Version        =   393216
               BackColorBkg    =   -2147483627
               FillStyle       =   1
               SelectionMode   =   1
               AllowUserResizing=   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
               _Band(0).GridLinesBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgCPCB 
               Height          =   945
               Left            =   0
               TabIndex        =   224
               Top             =   6330
               Width           =   15225
               _ExtentX        =   26855
               _ExtentY        =   1667
               _Version        =   393216
               BackColor       =   11927477
               BackColorBkg    =   -2147483627
               FillStyle       =   1
               SelectionMode   =   1
               AllowUserResizing=   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
               _Band(0).GridLinesBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin VB.Label Label53 
               Caption         =   "����"
               Height          =   195
               Left            =   8850
               TabIndex        =   226
               Top             =   6030
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Label54 
               Caption         =   "����"
               Height          =   285
               Left            =   11220
               TabIndex        =   225
               Top             =   6030
               Width           =   465
            End
         End
         Begin VB.Frame frmgc 
            Caption         =   "frmGC"
            Height          =   7275
            Index           =   2
            Left            =   -75240
            TabIndex        =   209
            Top             =   0
            Width           =   15555
            Begin VB.CommandButton cmdGG 
               Caption         =   "����"
               Height          =   315
               Left            =   13740
               TabIndex        =   213
               Top             =   5970
               Width           =   675
            End
            Begin VB.TextBox txtDj 
               Height          =   345
               Left            =   12120
               TabIndex        =   212
               Top             =   5970
               Width           =   1455
            End
            Begin VB.CommandButton cmdD 
               Caption         =   "ɾ��"
               Height          =   315
               Left            =   14490
               TabIndex        =   211
               Top             =   5970
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.TextBox txtTl 
               Height          =   315
               Left            =   9720
               TabIndex        =   210
               Top             =   5970
               Visible         =   0   'False
               Width           =   1515
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgBao 
               Height          =   5865
               Left            =   240
               TabIndex        =   214
               Top             =   0
               Width           =   15225
               _ExtentX        =   26855
               _ExtentY        =   10345
               _Version        =   393216
               BackColorBkg    =   -2147483627
               FillStyle       =   1
               SelectionMode   =   1
               AllowUserResizing=   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
               _Band(0).GridLinesBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgMa 
               Height          =   945
               Left            =   240
               TabIndex        =   215
               Top             =   6330
               Width           =   15225
               _ExtentX        =   26855
               _ExtentY        =   1667
               _Version        =   393216
               BackColor       =   11927477
               BackColorBkg    =   -2147483627
               FillStyle       =   1
               SelectionMode   =   1
               AllowUserResizing=   1
               _NumberOfBands  =   1
               _Band(0).Cols   =   2
               _Band(0).GridLinesBand=   1
               _Band(0).TextStyleBand=   0
               _Band(0).TextStyleHeader=   0
            End
            Begin VB.Label Label43 
               Caption         =   "����"
               Height          =   285
               Left            =   11460
               TabIndex        =   217
               Top             =   6030
               Width           =   465
            End
            Begin VB.Label Label42 
               Caption         =   "����"
               Height          =   195
               Left            =   9060
               TabIndex        =   216
               Top             =   6030
               Visible         =   0   'False
               Width           =   495
            End
         End
         Begin VB.Frame frmgc 
            Caption         =   "Frame4"
            Height          =   7305
            Index           =   1
            Left            =   -74970
            TabIndex        =   197
            Top             =   0
            Width           =   15195
            Begin VB.TextBox txtDxnr 
               Height          =   5385
               Left            =   0
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   208
               Top             =   1860
               Width           =   15195
            End
            Begin VB.Frame Frame5 
               Caption         =   "������Ϣ"
               ForeColor       =   &H000000FF&
               Height          =   1875
               Left            =   30
               TabIndex        =   198
               Top             =   180
               Width           =   15255
               Begin VB.Frame frmDx 
                  Height          =   375
                  Left            =   7170
                  TabIndex        =   200
                  Top             =   1170
                  Width           =   2235
                  Begin VB.TextBox txtMon 
                     Height          =   270
                     Left            =   1290
                     Locked          =   -1  'True
                     TabIndex        =   201
                     Top             =   120
                     Width           =   525
                  End
                  Begin VB.Label Label23 
                     Caption         =   "��"
                     Height          =   255
                     Left            =   1950
                     TabIndex        =   203
                     Top             =   120
                     Width           =   195
                  End
                  Begin VB.Label Label22 
                     Caption         =   "ά�ޱ�����"
                     DragMode        =   1  'Automatic
                     Height          =   225
                     Left            =   120
                     TabIndex        =   202
                     Top             =   120
                     Width           =   1065
                  End
               End
               Begin VB.TextBox txtZuD 
                  Height          =   285
                  Left            =   8430
                  Locked          =   -1  'True
                  TabIndex        =   199
                  Text            =   "Text1"
                  Top             =   615
                  Width           =   1725
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgB 
                  Height          =   1635
                  Left            =   0
                  TabIndex        =   204
                  Top             =   210
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
                  Left            =   8430
                  TabIndex        =   205
                  Top             =   210
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   582
                  _Version        =   393216
                  Locked          =   -1  'True
                  Text            =   "DataCombo2"
               End
               Begin VB.Label Label55 
                  Caption         =   "���̲����"
                  Height          =   225
                  Left            =   7230
                  TabIndex        =   207
                  Top             =   285
                  Width           =   945
               End
               Begin VB.Label Label56 
                  Caption         =   "�鳤"
                  Height          =   225
                  Left            =   7770
                  TabIndex        =   206
                  Top             =   675
                  Width           =   465
               End
            End
         End
         Begin VB.Frame frmgc 
            Caption         =   "Frame5"
            Height          =   7275
            Index           =   0
            Left            =   30
            TabIndex        =   170
            Top             =   30
            Width           =   15225
            Begin VB.Frame frmJi 
               Caption         =   "�ڲ��ɱ�"
               Height          =   2505
               Left            =   0
               TabIndex        =   192
               Top             =   3180
               Width           =   15195
               Begin VB.TextBox txtZu 
                  Height          =   285
                  Left            =   1440
                  Locked          =   -1  'True
                  TabIndex        =   193
                  Text            =   "Text1"
                  Top             =   750
                  Width           =   1725
               End
               Begin MSDataListLib.DataCombo comZu 
                  Height          =   330
                  Left            =   1440
                  TabIndex        =   194
                  Top             =   345
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   582
                  _Version        =   393216
                  Locked          =   -1  'True
                  Text            =   "DataCombo2"
               End
               Begin VB.Label Label36 
                  Caption         =   "�鳤"
                  Height          =   225
                  Left            =   690
                  TabIndex        =   196
                  Top             =   810
                  Width           =   465
               End
               Begin VB.Label Label35 
                  Caption         =   "���̲����"
                  Height          =   225
                  Left            =   150
                  TabIndex        =   195
                  Top             =   420
                  Width           =   945
               End
            End
            Begin VB.Frame Frame2 
               Caption         =   "������Ϣ"
               ForeColor       =   &H000000FF&
               Height          =   2145
               Left            =   0
               TabIndex        =   171
               Top             =   120
               Width           =   15195
               Begin VB.Frame frmNb 
                  BorderStyle     =   0  'None
                  Height          =   1815
                  Left            =   7560
                  TabIndex        =   172
                  Top             =   150
                  Width           =   7335
                  Begin VB.Frame frmTime 
                     Enabled         =   0   'False
                     Height          =   1665
                     Left            =   4290
                     TabIndex        =   177
                     Top             =   30
                     Width           =   3075
                     Begin VB.CheckBox chkBa 
                        Caption         =   "24Сʱ��ת"
                        Enabled         =   0   'False
                        Height          =   255
                        Left            =   270
                        TabIndex        =   180
                        Top             =   330
                        Width           =   1215
                     End
                     Begin VB.CheckBox chkBb 
                        Caption         =   "ȫ����ת"
                        Enabled         =   0   'False
                        Height          =   255
                        Left            =   270
                        TabIndex        =   179
                        Top             =   780
                        Width           =   1845
                     End
                     Begin VB.CheckBox chkBc 
                        Caption         =   "2Сʱ�ڵ���"
                        Enabled         =   0   'False
                        Height          =   255
                        Left            =   270
                        TabIndex        =   178
                        Top             =   1260
                        Width           =   1845
                     End
                     Begin VB.Label Label27 
                        Caption         =   "ʱ��ϵ��:"
                        Height          =   195
                        Left            =   300
                        TabIndex        =   181
                        Top             =   120
                        Width           =   1155
                     End
                  End
                  Begin VB.TextBox txtWc 
                     Height          =   270
                     Left            =   1050
                     Locked          =   -1  'True
                     TabIndex        =   176
                     Top             =   1440
                     Width           =   495
                  End
                  Begin VB.TextBox txtXc 
                     Height          =   270
                     Left            =   3330
                     Locked          =   -1  'True
                     TabIndex        =   175
                     Top             =   1440
                     Width           =   405
                  End
                  Begin VB.TextBox txtF 
                     Height          =   300
                     Left            =   30
                     Locked          =   -1  'True
                     TabIndex        =   174
                     Top             =   540
                     Width           =   1455
                  End
                  Begin VB.TextBox txtL 
                     Height          =   300
                     Left            =   2430
                     Locked          =   -1  'True
                     TabIndex        =   173
                     Top             =   540
                     Width           =   1305
                  End
                  Begin MSComCtl2.DTPicker dt4 
                     Height          =   315
                     Left            =   2430
                     TabIndex        =   182
                     Top             =   540
                     Width           =   1605
                     _ExtentX        =   2831
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   149094401
                     CurrentDate     =   38098
                  End
                  Begin MSComCtl2.DTPicker dt3 
                     Height          =   315
                     Left            =   30
                     TabIndex        =   183
                     Top             =   540
                     Width           =   1755
                     _ExtentX        =   3096
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   149094401
                     CurrentDate     =   38098
                  End
                  Begin VB.Label Label52 
                     Caption         =   "ά��������"
                     Height          =   225
                     Left            =   2550
                     TabIndex        =   190
                     Top             =   120
                     Width           =   1275
                  End
                  Begin VB.Label Label51 
                     Caption         =   "ά����ʼ��"
                     Height          =   225
                     Left            =   240
                     TabIndex        =   189
                     Top             =   150
                     Width           =   1605
                  End
                  Begin VB.Label Label16 
                     Caption         =   "ά������:"
                     Height          =   225
                     Left            =   60
                     TabIndex        =   188
                     Top             =   1470
                     Width           =   855
                  End
                  Begin VB.Label Label12 
                     Caption         =   "��"
                     Height          =   225
                     Left            =   1650
                     TabIndex        =   187
                     Top             =   1470
                     Width           =   255
                  End
                  Begin VB.Label Label10 
                     Caption         =   "�������"
                     Height          =   225
                     Left            =   2430
                     TabIndex        =   186
                     Top             =   1470
                     Width           =   825
                  End
                  Begin VB.Label Label21 
                     Caption         =   "��"
                     Height          =   225
                     Left            =   3840
                     TabIndex        =   185
                     Top             =   1470
                     Width           =   315
                  End
                  Begin VB.Label Label28 
                     Caption         =   "---��"
                     Height          =   225
                     Left            =   1950
                     TabIndex        =   184
                     Top             =   600
                     Width           =   375
                  End
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MMdtgA 
                  Height          =   1635
                  Left            =   30
                  TabIndex        =   191
                  Top             =   210
                  Width           =   6885
                  _ExtentX        =   12144
                  _ExtentY        =   2884
                  _Version        =   393216
                  SelectionMode   =   1
                  _NumberOfBands  =   1
                  _Band(0).Cols   =   2
               End
            End
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   30
            Left            =   -73200
            TabIndex        =   150
            Top             =   1530
            Width           =   30
         End
         Begin VB.Label Label14 
            Caption         =   "�ɹ��ɱ�"
            Height          =   225
            Left            =   -74880
            TabIndex        =   152
            Top             =   4050
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "����"
            Height          =   285
            Left            =   -63030
            TabIndex        =   151
            Top             =   3990
            Width           =   465
         End
      End
      Begin VB.Line Line3 
         X1              =   5040
         X2              =   5040
         Y1              =   3570
         Y2              =   7290
      End
      Begin VB.Line Line2 
         X1              =   5040
         X2              =   10350
         Y1              =   3570
         Y2              =   3570
      End
      Begin VB.Label Label30 
         Caption         =   "ǩ����"
         Height          =   255
         Left            =   7890
         TabIndex        =   107
         Top             =   1140
         Width           =   555
      End
      Begin VB.Label Label7 
         Caption         =   "��Ŀ����"
         Height          =   255
         Left            =   5250
         TabIndex        =   106
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "�ͻ�����"
         Height          =   255
         Left            =   240
         TabIndex        =   105
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label lblJlr 
         Caption         =   "�� �� 1"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   10680
         TabIndex        =   104
         Top             =   3840
         Width           =   915
      End
      Begin VB.Label Label26 
         Caption         =   "��Ŀ��ַ"
         Height          =   255
         Left            =   5250
         TabIndex        =   103
         Top             =   2190
         Width           =   885
      End
      Begin VB.Label Label24 
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
         Left            =   10680
         TabIndex        =   102
         Top             =   1110
         Width           =   885
      End
      Begin VB.Label Label17 
         Caption         =   "���ϳɱ�"
         Height          =   255
         Left            =   10680
         TabIndex        =   101
         Top             =   1565
         Width           =   825
      End
      Begin VB.Label Label9 
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
         Left            =   12000
         TabIndex        =   100
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "ά�ַ���"
         Height          =   255
         Left            =   10680
         TabIndex        =   99
         Top             =   3385
         Width           =   885
      End
      Begin VB.Label Label19 
         Caption         =   "��    ��"
         Height          =   255
         Left            =   10680
         TabIndex        =   98
         Top             =   180
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "�ְ����"
         Height          =   255
         Left            =   10680
         TabIndex        =   97
         Top             =   2930
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "�� �� ��"
         Height          =   255
         Left            =   10680
         TabIndex        =   96
         Top             =   2475
         Width           =   915
      End
      Begin VB.Label Label15 
         Caption         =   "�� �� ��"
         Height          =   255
         Left            =   10680
         TabIndex        =   95
         Top             =   2020
         Width           =   975
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
         Left            =   180
         TabIndex        =   94
         Top             =   3150
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "��ͬ����"
         Height          =   195
         Left            =   240
         TabIndex        =   93
         Top             =   2205
         Width           =   915
      End
      Begin VB.Label Label25 
         Caption         =   "��ͬ���"
         Height          =   225
         Left            =   240
         TabIndex        =   92
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "��    ��"
         Height          =   255
         Index           =   0
         Left            =   5250
         TabIndex        =   91
         Top             =   1665
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "��Ŀ������"
         Height          =   255
         Index           =   0
         Left            =   5100
         TabIndex        =   90
         Top             =   1170
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "�ͻ�����"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   89
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "��  ��"
         Height          =   255
         Left            =   8340
         TabIndex        =   88
         Top             =   1635
         Width           =   645
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "��   ͬ   ��   ��   ��"
         BeginProperty Font 
            Name            =   "���Ĳ���"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   4230
         TabIndex        =   87
         Top             =   120
         Width           =   2715
      End
      Begin VB.Label Label8 
         Caption         =   "�տ���"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8100
         TabIndex        =   86
         Top             =   3120
         Width           =   915
      End
      Begin VB.Label Label29 
         Caption         =   "�տ��ܶ�"
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   5250
         TabIndex        =   85
         Top             =   3150
         Width           =   795
      End
      Begin VB.Line Line1 
         X1              =   10350
         X2              =   10350
         Y1              =   7560
         Y2              =   0
      End
      Begin VB.Label lblHtxz 
         Height          =   315
         Left            =   1440
         TabIndex        =   84
         Top             =   2190
         Width           =   3315
      End
      Begin VB.Label Label49 
         Caption         =   "��ע"
         Height          =   225
         Left            =   5580
         TabIndex        =   83
         Top             =   2640
         Width           =   585
      End
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
      Height          =   345
      Left            =   6900
      TabIndex        =   237
      Top             =   8790
      Width           =   5475
   End
   Begin VB.Label lblMQM 
      Caption         =   "���ȷ��"
      Height          =   225
      Index           =   5
      Left            =   5970
      TabIndex        =   169
      Top             =   8100
      Width           =   885
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   5
      Left            =   5910
      TabIndex        =   168
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label lblMQM 
      Caption         =   "��ִͬ��"
      Height          =   225
      Index           =   4
      Left            =   4950
      TabIndex        =   166
      Top             =   8100
      Width           =   1185
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   4
      Left            =   4890
      TabIndex        =   165
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label lblMQM 
      Caption         =   "��ͬ����"
      Height          =   225
      Index           =   3
      Left            =   3960
      TabIndex        =   163
      Top             =   8100
      Width           =   1185
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   3
      Left            =   3900
      TabIndex        =   162
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label lblMQM 
      Caption         =   "�����ܼ�"
      Height          =   225
      Index           =   2
      Left            =   2970
      TabIndex        =   160
      Top             =   8100
      Width           =   1185
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   2
      Left            =   2910
      TabIndex        =   159
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label lblMQM 
      Caption         =   "���۾���"
      Height          =   225
      Index           =   1
      Left            =   1950
      TabIndex        =   157
      Top             =   8100
      Width           =   1185
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   1890
      TabIndex        =   156
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label lblJiLI 
      Caption         =   "���ٴΰ��ύ��ť,�Ա�ˢ������"
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   13080
      TabIndex        =   125
      Top             =   8160
      Width           =   1725
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   870
      TabIndex        =   115
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label lblMQM 
      Caption         =   "ҵ��Ա"
      Height          =   225
      Index           =   0
      Left            =   930
      TabIndex        =   114
      Top             =   8100
      Width           =   1185
   End
   Begin VB.Label lblMHid 
      Caption         =   "lblHid"
      Height          =   285
      Left            =   7410
      TabIndex        =   113
      Top             =   8370
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblBaoid 
      Caption         =   "lblBaoid"
      Height          =   285
      Left            =   9030
      TabIndex        =   112
      Top             =   8250
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "frmMNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public adoOid As New ADODB.Recordset '����Old���ӵ�ADO
'Public adoBx As ADODB.Recordset '�ɹ���(���)
'Public adoGx As ADODB.Recordset '�ɱ���(���)
'Public adoBxCP As ADODB.Recordset '�ɹ���(��Ʒ)
'Public adoGxCP As ADODB.Recordset '�ɱ���(��Ʒ)
'Public adoFFk As ADODB.Recordset 'Ԥ�Ƹ���
'Public adoYj As ADODB.Recordset '�ʽ��



'Public adoA As ADODB.Recordset
'Public adoB As ADODB.Recordset

Dim timZm As Integer '�����ύ��,��timWaitִ�еĺ�������ID(2 �����ͬ 3�½�ѯ�۵�(���),6�½�ѯ�۵�(��Ʒ),10ǩ��11���ɺ�ͬ���)

Private Sub chkD_Click()
If chkC.Value = 1 Then
    tabHt.Tab = 1
    tabGc.TabVisible(5) = True
    
End If
End Sub


Private Sub cmdAdd_Click()
On Error Resume Next
Set mod1.cmd = New ADODB.command
mod1.cmd.ActiveConnection = mod1.CC
mod1.cmd.CommandText = "htFkAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@rq") = txtYrq.Text
mod1.cmd.Parameters("@yingfJe") = Round(Val(txtHtze.Text) * Val(txtYed.Text) / 100, 2)
mod1.cmd.Parameters("@htbh") = lblMHid.Caption
mod1.cmd.Parameters("@ed") = Round(Val(txtYed.Text) / 100, 2)
mod1.cmd.Execute
Set cmd = Nothing

txtYed.Text = ""
mod1.mFk.Requery
Set MMdtgFk.DataSource = mod1.mFk
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
If htBrow.Visible = True Then
    htBrow.Enabled = True
    htBrow.ZOrder 0
ElseIf htBrowG.Visible = True Then
    htBrowG.Enabled = True
    htBrowG.ZOrder 0
ElseIf Dialog.Visible = True Then
    Dialog.ZOrder 0
    Dialog.Enabled = True
End If

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
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "baoJiaGx"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@dj") = Val(txtCj.Text)
    mod1.cmd.Parameters("@sl") = Val(txtCL.Text)
    mod1.cmd.Parameters("@lid") = liD
    mod1.cmd.Execute
    'txtHg.Text = Val(txtHg.Text) + mod1.CMD.Parameters("@hg").Value
    Set cmd = Nothing
    
'    tt = "select bid from baojiaD where baoid=" & Val(lblBaoid.Caption)
'    Set mod1.HTP = New ADODB.Recordset
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    bid = mod1.HTP.Fields("bid").Value
'    If lblHtxz.Caption = "ά��" Or lblHtxz.Caption = "����" Then
'        '�����Ӧѯ�۵���cgid��
'        tt = "select cgid from xunJiaD where bid=" & bid
'        Set mod1.HTP = New ADODB.Recordset
'        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'        bid = mod1.HTP.Fields("cgid").Value
'    End If
'
'    '������Ӧѯ����ϸ�е�����
'    tt = "update XunJiaMx set sl=" & Val(txtTl.Text) & ",hg=dj*" & Val(txtTl.Text) & " where lid=" & liD
'    Set mod1.HTP = New ADODB.Recordset
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'    '������Ӧѯ�۵��еĽ��
'    tt = "select sum(hg) as hg from xunjiamx where bid=" & bid
'    Set mod1.HTP = New ADODB.Recordset
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''    XCB = 0
''    Do While Not mod1.HTP.EOF
''        XCB = XCB + mod1.HTP.Fields("hg").Value
''        mod1.HTP.MoveNext
''    Loop
'    XCB = mod1.HTP.Fields("hg").Value
'
'    tt = "update xunjiaD set hg=" & XCB & ",yhg=" & XCB & " where bid=" & bid
'    Set mod1.HTP = New ADODB.Recordset
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
frmYM.Visible = False
End Sub

Private Sub cmdCong_Click()
Dim tt As String
Dim Bid As Long
Dim ZL As String
Dim oo As Integer
Dim ii As Integer
On Error Resume Next
If Not (optZ.Value = True Or optW.Value = True) Then
    If mod1.DName = lblYwy.Caption Then





        ii = MsgBox("�������������ʹԭ�ȵ�������ִ�е�����ȫ������,�Ƿ�ȷ��ִ��?", vbYesNo + vbInformation, "ѯ��")
        If ii = vbYes Then
            tt = InputBox("��������Ҫ���ص�ԭ��!")
            If tt = "" Then Exit Sub
            Set mod1.cmd = New ADODB.command
            mod1.cmd.ActiveConnection = mod1.CC
            mod1.cmd.CommandText = "xtzxFAdd"
            mod1.cmd.CommandType = adCmdStoredProc
            mod1.cmd.Parameters("@yid").Value = 62 '��ǩ��
            mod1.cmd.Parameters("@lc").Value = 2 '�˻����������
            mod1.cmd.Parameters("@bh").Value = txtHtbh.Text
            mod1.cmd.Parameters("@ywy").Value = mod1.DName
            mod1.cmd.Parameters("@uid").Value = mod1.DHid
            mod1.cmd.Parameters("@bz").Value = tt
            mod1.cmd.Parameters("@zn").Value = "new" '���ְ��
            mod1.cmd.Execute
            Set cmd = Nothing
            For oo = 0 To 6
                cmdMQm(oo).Caption = ""
                lblMTm(oo).Caption = ""
            Next
            lblLc.Caption = 999 '�����ٰ�ǩ����ť.
            If Dialog.Visible = True Then '���������б�
                Call mod1.refEnvent
            End If
            Exit Sub
        End If
    End If
Else
    MsgBox ("��ͬ�Ѿ���ʽ����,�����޸�!")
End If
End Sub

Private Sub cmdDe_Click()
Dim tt As String
On Error Resume Next
tt = "delete from htping1 where fid=" & Val(lblFid.Caption)
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

txtYed.Text = ""
mod1.mFk.Requery
Set MMdtgFk.DataSource = mod1.mFk
End Sub

Private Sub cmdGB_Click()
frmWai.Visible = False

End Sub

Private Sub cmdDing_Click()
Dim tt As String
On Error Resume Next

If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("����һ��Ҫ���߾ܾ��ҵ�����!  :) ")
    Exit Sub
End If
timZm = 10 'ǩ��
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
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
    mod1.cmd.Parameters("@mt3") = txtXMMC.Text
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
    If OptT1.Value = True Then
        mod1.cmd.Parameters("@mb1") = 1 'ͬ��
    Else
        mod1.cmd.Parameters("@mb1") = 0 '�ܾ�
    End If
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = Null
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    
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
End Sub


Private Sub cmdGG_Click()
Dim CB As Long
Dim liD As Long
Dim Bid As Long
Dim XCB As Long
On Error Resume Next
If Val(txtDj.Text) = 0 Then Exit Sub
MMdtgBao.Col = 16
liD = MMdtgBao.Text
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "baoJiaGx"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@dj") = Val(txtDj.Text)
    mod1.cmd.Parameters("@sl") = Val(txtTl.Text)
    mod1.cmd.Parameters("@lid") = liD
    mod1.cmd.Execute
    'txtHg.Text = Val(txtHg.Text) + mod1.CMD.Parameters("@hg").Value
    Set cmd = Nothing
    
'    tt = "select bid from baojiaD where baoid=" & Val(lblBaoid.Caption)
'    Set mod1.HTP = New ADODB.Recordset
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    bid = mod1.HTP.Fields("bid").Value
'    If lblHtxz.Caption = "ά��" Or lblHtxz.Caption = "����" Then
'        '�����Ӧѯ�۵���cgid��
'        tt = "select cgid from xunJiaD where bid=" & bid
'        Set mod1.HTP = New ADODB.Recordset
'        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'        bid = mod1.HTP.Fields("cgid").Value
'    End If
'
'    '������Ӧѯ����ϸ�е�����
'    tt = "update XunJiaMx set sl=" & Val(txtTl.Text) & ",hg=dj*" & Val(txtTl.Text) & " where lid=" & liD
'    Set mod1.HTP = New ADODB.Recordset
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'    '������Ӧѯ�۵��еĽ��
'    tt = "select sum(hg) as hg from xunjiamx where bid=" & bid
'    Set mod1.HTP = New ADODB.Recordset
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''    XCB = 0
''    Do While Not mod1.HTP.EOF
''        XCB = XCB + mod1.HTP.Fields("hg").Value
''        mod1.HTP.MoveNext
''    Loop
'    XCB = mod1.HTP.Fields("hg").Value
'
'    tt = "update xunjiaD set hg=" & XCB & ",yhg=" & XCB & " where bid=" & bid
'    Set mod1.HTP = New ADODB.Recordset
'    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    txtDj.Text = ""
    txtSL.Text = ""
   ' txtClcb.Text = XCB
    mod1.mBx.Requery
    Set MMdtgBao.DataSource = mod1.mBx
   ' Call cmdSave_Click
    txtDj.Text = ""
    txtTl.Text = ""
End Sub

Private Sub cmdGx_Click()
On Error Resume Next
Set mod1.cmd = New ADODB.command
mod1.cmd.ActiveConnection = mod1.CC
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
Dim Qy As String
Dim xZ As String
Dim XZDm As String
'�жϺ�ͬ���ʺͺ�ͬ���.
If Val(txtH1.Text) > 0 Then
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
End If
    txtHtbh.Text = "HM" & Qy & Format(mod1.DQda, "yyyymmdd") & XZDm & lblMHid.Caption
    lblHtxz.Caption = xZ
    
timZm = 11 '���ɺ�ͬ���
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
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

End Sub

Private Sub cmdMod_Click()
On Error Resume Next
''If lblLcUid.Caption <> mod1.DHid Then
''Exit Sub
''End If
cmdYadd.Visible = False
cmdYdel.Visible = False
If (lblLc.Caption = 1 Or lblLc.Caption = 0) And txtYwy.Text = mod1.DName Then
    frmFX.Visible = True
    dt3.Enabled = True
    dt4.Enabled = True
    optLa.Enabled = True
    optLb.Enabled = True
    optLc.Enabled = True
    cmdSave.Enabled = True
    txtHtze.Locked = False
ElseIf mod1.BmJl = True Then
    frmFX.Visible = True
    dt3.Enabled = True
    dt4.Enabled = True
    optLa.Enabled = True
    optLb.Enabled = True
    optLc.Enabled = True
    cmdSave.Enabled = True
    txtHtze.Locked = False
ElseIf (mod1.DName = "����" Or mod1.DName = "������" Or mod1.DName = "������1" Or mod1.DName = "������") And optW.Value = False Then
    frmPL.Visible = True
    frmFX.Visible = True
    dt3.Enabled = True
    dt4.Enabled = True
    optLa.Enabled = True
    optLb.Enabled = True
    optLc.Enabled = True
    'txtYf1.Locked = False
    txtQt1.Locked = False
    txtYj1.Locked = False
    txtTcBe.Locked = False
    txtHtze.Locked = False
    cmdSave.Enabled = True
    txtFbje1.Locked = False
    If lblyjFF.Caption = "False" Then
        cmdYadd.Visible = True
        cmdYdel.Visible = True
    End If
    'JILI = 0
End If
End Sub

Private Sub cmdPje_Click()
Dim tt As String
On Error Resume Next
Pje.Show
Set Pje.adoPje = New ADODB.Recordset
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



Dim tt As String
Dim Tywy As String '������ת����һ�˵�����
Dim Tuid As String
Dim Oywy As String 'ԭ����ת�˵�����
Dim Ouid As String 'ԭ����ת�˵Ĺ���

On Error Resume Next
If cmdMQm(Index).Caption <> "" Then
    Exit Sub
End If
If mod1.mFk.RecordCount = 0 Then
    MsgBox ("�����븶�ʽ!")
    cmdSave.Enabled = True
    Exit Sub
End If

If optLa.Value = False And optLb.Value = False And optLc.Value = False Then
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

If cmdSave.Enabled = True Then
    MsgBox "���Ƚ����ӱ���,��ǩ�����Ĵ���!"
    Exit Sub
End If

If Index + 1 <> lblLc.Caption And lblLc.Caption <> 0 Then '�����ڲ���ɵ�λ�����ҵ�
    Exit Sub
End If

If lblLcUid.Caption <> mod1.DHid Then
    MsgBox "�˴�Ӧ��" & lblLcRen.Caption & "ǩ��! ������Ҫ�ٵ�"
    Exit Sub
End If


If txtHtbh.Text = "" Then
    MsgBox ("�������ɺ�ͬ���!")
    Exit Sub
End If
frmQm.Visible = True
Exit Sub





If lblLc.Caption > 1 Then
    ii = MsgBox("���Ƿ��׼�˵���(ѡ���ǡ���ǩ��ͨ��,ѡ�񡰷񡱽����ش˵�)", vbYesNoCancel + vbInformation, "����ע��!")
    If ii = vbNo Then
        ii = MsgBox("�����ص����۵��ĳ�ʼ����!", vbYesNo + vbInformation, "ȷ�ϲ�����?")
        If ii = vbNo Then
            Exit Sub
        End If
        tt = InputBox("��������Ҫ���ص�ԭ��!")
        Set mod1.cmd = New ADODB.command
        mod1.cmd.ActiveConnection = mod1.CC
        mod1.cmd.CommandText = "xtzxFAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@yid").Value = 62 '��ǩ��
        mod1.cmd.Parameters("@lc").Value = lblLc.Caption
        mod1.cmd.Parameters("@bh").Value = txtHtbh.Text
        mod1.cmd.Parameters("@ywy").Value = mod1.DName
        mod1.cmd.Parameters("@uid").Value = mod1.DHid
        mod1.cmd.Parameters("@bz").Value = tt
        mod1.cmd.Parameters("@zn").Value = lblMQM(Index).Caption '���ְ��
        mod1.cmd.Execute
        Set cmd = Nothing
        For oo = 0 To 5
            cmdMQm(oo).Caption = ""
            lblMTm(oo).Caption = ""
        Next
        lblLc.Caption = 999 '�����ٰ�ǩ����ť.
        If Dialog.Visible = True Then '���������б�
            Call mod1.refEnvent
        End If
        Exit Sub
    ElseIf ii = vbCancel Then
        Exit Sub
    End If
ElseIf lblLc.Caption = 0 Then
    Dim Zi As Integer
    Zi = MsgBox("�Ƿ�ȷ��ǩ��?", vbYesNo)
    If Zi = vbNo Then Exit Sub
End If



Oywy = lblLcRen.Caption
Ouid = lblLcUid.Caption

If lblLc.Caption = 0 Or lblLc.Caption = 1 Then
Dim Qy As String
Dim xZ As String
Dim XZDm As String
'�жϺ�ͬ���ʺͺ�ͬ���.
If Val(txtH1.Text) > 0 Then
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
ElseIf Val(txtW5.Text) > 0 Then
    xZ = "�����"
    XZDm = "LP"
ElseIf Val(txtW6.Text) > 0 Then
    xZ = "��Ʒ"
    XZDm = "CP"
End If
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
End If
    txtHtbh.Text = "HM" & Qy & Format(mod1.DQda, "yyyymmdd") & XZDm & lblMHid.Caption
    lblHtxz.Caption = xZ
    '���ǩ�ֱ�qmrz
    tt = "insert into QMrz (Qlabel,Zid,btz,QDBh)  select Lnr,zid,6," & lblMHid.Caption & "  from NewLCMX where yid=80"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    tt = "update htping set htbh='" & txtHtbh.Text & "',htxz='" & xZ & "' where hid=" & lblMHid.Caption
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    tt = "select username,userid from worker where zzf=1 and bm='" & mod1.BM & "' and bmjl=1"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Tywy = mod1.HTP.Fields("username").Value
    Tuid = mod1.HTP.Fields("userid").Value

lblLc.Caption = 1
If Val(txtHtze.Text) < 15000 And lblHtxz.Caption <> "ά��" And lblHtxz.Caption <> "����" Then
    cmdMQm(2).Enabled = False
    
End If
End If

    

    lblLc.Caption = lblLc.Caption + 1
If Val(txtHtze.Text) < 15000 And lblHtxz.Caption <> "ά��" And lblHtxz.Caption <> "����" Then
    lblLc.Caption = lblLc.Caption + 1
    
End If
    If lblLc.Caption = 3 Then
        Tywy = "����"
        Tuid = "HM040"
    ElseIf lblMQM(Index + 1) = "�������" Then
        If comQy.Text = "�Ϻ�" Then

        ElseIf comQy.Text = "�Ͼ�" Then
            Tywy = "����"
            Tuid = "HM051"
        ElseIf comQy.Text = "����" Then
            Tywy = "����"
            Tuid = "HM316"
        ElseIf comQy.Text = "����" Then
            Tywy = "����֥"
            Tuid = "HM190"
        ElseIf comQy.Text = "����" Then
            Tywy = "����"
            Tuid = "HMG010"
        End If
        tt = "update htping set htf=9 where hid=" & Val(lblMHid.Caption)
        Set mod1.HTP = New ADODB.Recordset
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    ElseIf lblMQM(Index + 1) = "��ִͬ��" Then
        If comQy.Text = "�Ϻ�" Then
            Tywy = "���"
            Tuid = "HM233"
        ElseIf comQy.Text = "�Ͼ�" Then
            Tywy = "����"
            Tuid = "HM051"
        ElseIf comQy.Text = "����" Then
            Tywy = "����"
            Tuid = "HM316"
        ElseIf comQy.Text = "����" Then
            Tywy = "����֥"
            Tuid = "HM190"
        ElseIf comQy.Text = "����" Then
            Tywy = "����"
            Tuid = "HMG010"
        End If
        tt = "update htping set htf=1,htrq='" & Date & "' where hid=" & Val(lblMHid.Caption)
        Set mod1.HTP = New ADODB.Recordset
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    ElseIf lblLc.Caption = 6 Then
        tt = "update htping set htf=2 where hid=" & Val(lblMHid.Caption)
        Set mod1.HTP = New ADODB.Recordset
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    End If
    
''''''If lblmqm(Index).Caption = "��ִͬ��" Then
''''''        Set mod1.CMD = New ADODB.command
''''''        mod1.CMD.ActiveConnection = mod1.CC
''''''        mod1.CMD.CommandText = "TXht"
''''''        mod1.CMD.CommandType = adCmdStoredProc
''''''        mod1.CMD.Parameters("@lc") = Val(lblLc.Caption)
''''''        mod1.CMD.Parameters("@htbh") = txtHtbh.Text
''''''        mod1.CMD.Parameters("@fwid") = Val(lblFwid.Caption)
''''''        mod1.CMD.Parameters("@nr") = txtXMMC.Text
''''''        mod1.CMD.Parameters("@khdh") = txtKhdm.Text
''''''        mod1.CMD.Parameters("@uid") = lblUid.Caption
''''''        mod1.CMD.Parameters("@bm") = lblBm.Caption
''''''        mod1.CMD.Parameters("@Errch") = ""   '������
''''''        mod1.CMD.Execute
''''''        If mod1.CMD.Parameters("@Errch").Value <> "�ɹ�" Then
''''''        lblLc.Caption = lblLc.Caption - 1
''''''        MsgBox "������ֹ���,������һ��,��������ύ���ɹ�,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
''''''        Exit Sub
''''''        End If
''''''    cmdmqm(Index).Caption = mod1.DName
''''''    lblmtm(Index).Caption = mod1.DQda
''''''    lblLcRen.Caption = ""
''''''    lblLcUid.Caption = ""

''''''End If
''''''    MsgBox ("���ݵ���ɹ�!������,���������������˵���ִ��!")
''''''Exit Sub
''''''End If

''''''''''''''If lblLc.Caption <> 2 Then
''''''''''''''
''''''''''''''    '���±�baojiaD�е�lcRen,lcUid �ֶ�,�Լ�QMRZ���е���Ӧ�ֶ�.
''''''''''''''                Set mod1.cmd = New ADODB.command
''''''''''''''                mod1.cmd.ActiveConnection = mod1.CC
''''''''''''''                mod1.cmd.CommandText = "QMRZQM"
''''''''''''''                mod1.cmd.CommandType = adCmdStoredProc
''''''''''''''                mod1.cmd.Parameters("@nlb") = lblNlb.Caption '����(������)����
''''''''''''''                mod1.cmd.Parameters("@lc") = lblLc.Caption  '��ǰ����
''''''''''''''                mod1.cmd.Parameters("@Dname") = mod1.DName
''''''''''''''                mod1.cmd.Parameters("@uid") = mod1.DHid
''''''''''''''                mod1.cmd.Parameters("@btz") = mod1.BTZ 'ҵ������
''''''''''''''                mod1.cmd.Parameters("@zid") = Index + 1 '����˳��
''''''''''''''                mod1.cmd.Parameters("@Qdbh") = Trim(txtHtbh.Text)   '���ӱ��
''''''''''''''                mod1.cmd.Parameters("@pje") = ""   '������
''''''''''''''                mod1.cmd.Parameters("@bm") = ""
''''''''''''''                mod1.cmd.Parameters("@qy") = ""
''''''''''''''                mod1.cmd.Parameters("@Gren") = "" '���Ϊ���ù���������,����ӷ��ù����˵Ĳ���
''''''''''''''                mod1.cmd.Parameters("@Guid") = ""
''''''''''''''                mod1.cmd.Parameters("@ywy") = lblYwy.Caption
''''''''''''''                mod1.cmd.Parameters("@yid") = lblUid.Caption
''''''''''''''                mod1.cmd.Parameters("@comid") = mod1.comId
''''''''''''''                mod1.cmd.Execute
''''''''''''''                Tywy = mod1.cmd.Parameters("@Tywy").Value
''''''''''''''                Tuid = mod1.cmd.Parameters("@Tuid").Value
''''''''''''''                Set cmd = Nothing
''''''''''''''                cmdmqm(Index).Caption = mod1.DName
''''''''''''''                lblmtm(Index).Caption = mod1.DQda
''''''''''''''
''''''''''''''Else
''''''''''''''
''''''''''''''    If mod1.comId = 0 And Not (mod1.Bm = "ά����3" Or mod1.Bm = "��Ʒ��1" Or mod1.Bm = "��Ʒ��2") Then
''''''''''''''        Tywy = "����"
''''''''''''''        Tuid = "HM040"
''''''''''''''    Else
''''''''''''''        If mod1.comId = 0 Then
''''''''''''''            Tywy = "������"
''''''''''''''            Tuid = "HM003"
''''''''''''''        ElseIf mod1.comId = 1 Then
''''''''''''''            Tywy = "������1"
''''''''''''''            Tuid = "HMG000"
''''''''''''''        End If
''''''''''''''    End If



    tt = "update QMRZ set  Qren='" & mod1.DName & "',Qrid='" & mod1.DHid & "',Qrq='" & mod1.DQda & "',xf=1 where Qdbh='" & lblMHid.Caption & "' and btz=6 and zid=" & (Index + 1)
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    tt = "update htping set lcren='" & Tywy & "',lcUid='" & Tuid & "',lc=" & Val(lblLc.Caption) & " where hid=" & lblMHid.Caption
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    
    cmdMQm(Index).Caption = mod1.DName
    lblMTm(Index).Caption = mod1.DQda
'End If


If lblMQM(Index + 1).Caption = "�������" Then
    If comQy.Text = "�Ϻ�" Then

    ElseIf comQy.Text = "�Ͼ�" Then
        Tywy = "����"
        Tuid = "HM051"
    ElseIf comQy.Text = "����" Then
        Tywy = "����"
        Tuid = "HM316"
    ElseIf comQy.Text = "����" Then
        Tywy = "����֥"
        Tuid = "HM190"
    ElseIf comQy.Text = "����" Then
        Tywy = "����"
        Tuid = "HMG010"
    End If
    tt = "update htping set lcren='" & Tywy & "',lcUid='" & Tuid & "',lc=" & Val(lblLc.Caption) & " where htbh='" & txtHtbh.Text & "'"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
ElseIf lblMQM(Index + 1).Caption = "��ִͬ��" Then
    If comQy.Text = "�Ϻ�" Then
        Tywy = "���"
        Tuid = "HM233"
    ElseIf comQy.Text = "�Ͼ�" Then
        Tywy = "����"
        Tuid = "HM051"
    ElseIf comQy.Text = "����" Then
        Tywy = "����"
        Tuid = "HM316"
    ElseIf comQy.Text = "����" Then
        Tywy = "����֥"
        Tuid = "HM190"
    ElseIf comQy.Text = "����" Then
        Tywy = "����"
        Tuid = "HMG010"
    End If
    tt = "update htping set lcren='" & Tywy & "',lcUid='" & Tuid & "',lc=" & Val(lblLc.Caption) & " where htbh='" & txtHtbh.Text & "'"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
End If
lblLcRen.Caption = Tywy
lblLcUid.Caption = Tuid

'''''''''''''''''lblLcRen.Caption = Tywy
'''''''''''''''''lblLcUid.Caption = Tuid

''''''''''''''''''''Select Case lblmqm(Index).Caption
''''''''''''''''''''Case "�������"
''''''''''''''''''''    tt = "update htping set htf=9 where hid=" & Val(lblmhid.Caption)
''''''''''''''''''''    Set mod1.HTP = New ADODB.Recordset
''''''''''''''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
''''''''''''''''''''Case "��ִͬ��"
''''''''''''''''''''    tt = "update htping set htf=1,htrq='" & Date & "' where hid=" & Val(lblmhid.Caption)
''''''''''''''''''''    Set mod1.HTP = New ADODB.Recordset
''''''''''''''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
''''''''''''''''''''Case "ִ�����ȷ��"
''''''''''''''''''''    tt = "update htping set htf=2 where hid=" & Val(lblmhid.Caption)
''''''''''''''''''''    Set mod1.HTP = New ADODB.Recordset
''''''''''''''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
''''''''''''''''''''End Select
''''''''''''''''''''
''''''''''''''''''''If Val(lblLc.Caption) > Val(lblLcou.Caption) And lblLc.Caption <> 1 Then
''''''''''''''''''''    Call mod1.EnventFinish(frmWbNew.lblFwid.Caption)
''''''''''''''''''''    tt = "update htping set Pwf=1 where hid=" & Val(lblmhid.Caption)
''''''''''''''''''''    Set mod1.HTP = New ADODB.Recordset
''''''''''''''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
''''''''''''''''''''    MsgBox "���������ݺ�ͬ��!"
''''''''''''''''''''
''''''''''''''''''''
''''''''''''''''''''Else
'''''''''''''''''''''    If lblLc.Caption = 1 Then 'ҵ��Ա��һ��ǩ��,��ѯ�����ڵ���ǩ������
'
'    End If
    '�������
    If lblLc.Caption <> 6 Then
        Call mod1.EnventAdd("��ͬ����", txtXMMC.Text, lblLcRen.Caption, lblLcUid.Caption, lblMHid.Caption, lblMQM(Index + 1).Caption, Oywy, Ouid, txtYwy.Text, txtYwy.ToolTipText, Val(lblFwid.Caption), lblMHid.Caption)
    End If
    Select Case lblMQM(Val(lblLc.Caption) - 1).Caption
    Case "�������"
        MsgBox "���ȫ��ͨ��,�˵�����ͬ�ͻ�������!"
    Case "��ִͬ��"
        MsgBox "����,��ѯ�۵������� " & Tywy & " ������!"
    Case "ִ�����ȷ��"
        MsgBox "������Ϣ������" & lblYwy.Caption & "ȥע����ݺ�ͬ!"
    Case Else
        MsgBox "����,��ѯ�۵������� " & Tywy & " ������!"
    End Select
    

'''''''End If

timZm = 10 'ǩ��
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "ǩ��"
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
    mod1.cmd.Parameters("@mlt1") = txtBz.Text '������
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = lblLc.Caption
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
    End If

    
Set mod1.cmd = Nothing

End Sub

Private Sub cmdSave_Click()
Dim W1 As Single
Dim W2 As Single
Dim W3 As Single
Dim W5 As Single
Dim W6 As Single
Dim FPLX As String

'����˼۳ɱ�����Ԥ�Ƴɱ�,��ȡ�˼۳ɱ�,�����ȡԤ�Ƴɱ��������.

    W1 = Val(txtH1.Text)


    W2 = Val(txtH2.Text)


'If Val(txtH3.Text) > Val(txtW3.Text) Then
'    W3 = Val(txtH3.Text)
'Else
    W3 = Val(txtW3.Text)
'End If
If Val(txtH5.Text) > Val(txtW5.Text) Then
    W5 = Val(txtH5.Text)
Else
    W5 = Val(txtW5.Text)
End If
If Val(txtH6.Text) > Val(txtW6.Text) Then
    W6 = Val(txtH6.Text)
Else
    W6 = Val(txtW6.Text)
End If

txtRgf1.Text = W1 + W2
txtFbje1.Text = W3 + Val(txtW4.Text)
txtClcb1.Text = W5 + W6

If lblHtxz.Caption = "ά��" Or lblHtxz.Caption = "ά��" Then
'����ɱ�����
    txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + Val(txtYf1.Text)
    txtJlr1.Text = Val(txtHtze.Text) - Val(txtCbze1.Text)
    txtLr1.Text = Val(txtJlr1.Text) - Val(txtYj1.Text)
    txtQt1.Text = Val(txtLr1.Text) * 0.1
    
    txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + Val(txtYf1.Text) + Val(txtQt1.Text)
    txtJlr1.Text = Val(txtHtze.Text) - Val(txtCbze1.Text)
    txtLr1.Text = Val(txtJlr1.Text) - Val(txtYj1.Text)
Else
    txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + Val(txtYf1.Text) + Val(txtQt1.Text)
    txtJlr1.Text = Val(txtHtze.Text) - Val(txtCbze1.Text)
    txtLr1.Text = Val(txtJlr1.Text) - Val(txtYj1.Text)
End If

If optLa.Value = True Then
    FPLX = "��ֵ��Ʊ"
ElseIf optLb.Value = True Then
    FPLX = "��ҵ��Ʊ"
ElseIf optLc.Value = True Then
    FPLX = "����Ʊ"
End If
If txtTcRQ.Text = "" Then
    txtTcRQ.Text = "2000-1-1"
End If




timZm = 2 '�����ͬ
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
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
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = txtBz.Text '��ע
    mod1.cmd.Parameters("@mlt2") = txtWBNR.Text '�������
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtYf1.Text) '�˷�
    mod1.cmd.Parameters("@mm2") = Val(txtTcBe.Text) '��ɱ���
    mod1.cmd.Parameters("@mm3") = Val(lblLc.Caption) '�������Ϊ0,�����ҵ��Ա������
    mod1.cmd.Parameters("@mm4") = 0
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
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = FMXC.dt3.Value 'ά����ʼ��
    mod1.cmd.Parameters("@md2") = FMXC.dt4.Value
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    
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
    End If

    
Set mod1.cmd = Nothing

End Sub

Private Sub cmdW1_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If Val(cmdW1.ToolTipText) > 0 Then
mod1.BTZ = 36
Call modBJD.BJDWBQing
Call modBJD.BJDBound(cmdW1.ToolTipText, "ά��")
frmWBXJ.Show
Exit Sub
End If
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
Set mod1.cmd = New ADODB.command
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
mod1.cmd.Parameters("@xmmc") = txtXMMC.Text
mod1.cmd.Parameters("@xid") = txtXMMC.ToolTipText
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
frmWBXJ.lblBm.Caption = mod1.BM
frmWBXJ.lblQy.Caption = mod1.Qy
frmWBXJ.lblZl.Caption = "ά��"
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
Set mod1.mA = New ADODB.Recordset
mod1.mA.Close
mod1.mA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.dtgA.DataSource = mod1.mA


'���º�ͬ
tt = "update htping set bid1=" & Val(frmWBXJ.lblBid.Caption) & "where hid=" & Val(lblMHid.Caption)
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
cmdW1.ToolTipText = frmWBXJ.lblBid.Caption


frmWBXJ.cmdBjd.Visible = False
frmWBXJ.txtHg.Locked = True
frmWBXJ.txtYhg.Locked = True
frmWBXJ.txtClf.Locked = True
frmWBXJ.cmdCg.Enabled = False
'frmWBXJ.cmdCong.Visible = False
frmWBXJ.cmdTk.Visible = True
frmWBXJ.Visible = True
frmWBXJ.comXmmc.Text = txtXMMC.Text
frmWBXJ.comXmmc.ToolTipText = txtXMMC.ToolTipText
frmWBXJ.cmdSave.Enabled = True

End Sub


Private Sub cmdW2_Click()
'Call modBJD.BJDWBQing
'frmWBXJ.Visible = True

Dim tt As String
Dim ii As Integer
On Error Resume Next
If Val(cmdW2.ToolTipText) > 0 Then
mod1.BTZ = 36
Call modBJD.BJDWBQing
Call modBJD.BJDBound(cmdW2.ToolTipText, "����")
frmWBXJ.Show
frmWBXJ.cmdSave.Enabled = True
frmWBXJ.frmTime.Visible = False
frmWBXJ.frmNb.Visible = False
frmWBXJ.cmdD.Visible = False
frmWBXJ.cmdTk.Visible = False
frmWBXJ.cmdCg.Visible = False
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
Set mod1.cmd = New ADODB.command
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
mod1.cmd.Parameters("@xmmc") = txtXMMC.Text
mod1.cmd.Parameters("@xid") = txtXMMC.ToolTipText
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
frmWBXJ.lblBm.Caption = mod1.BM
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
Set mod1.mA = New ADODB.Recordset
mod1.mA.Close
mod1.mA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.dtgA.DataSource = mod1.mA


'���º�ͬ
tt = "update htping set bid2=" & Val(frmWBXJ.lblBid.Caption) & "where hid=" & Val(lblMHid.Caption)
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
cmdW2.ToolTipText = frmWBXJ.lblBid.Caption


frmWBXJ.cmdBjd.Visible = False
frmWBXJ.txtHg.Locked = True
frmWBXJ.txtYhg.Locked = True
frmWBXJ.txtClf.Locked = True
frmWBXJ.cmdCg.Enabled = False
'frmWBXJ.cmdCong.Visible = False
frmWBXJ.cmdTk.Visible = True
frmWBXJ.Visible = True
frmWBXJ.comXmmc.Text = txtXMMC.Text
frmWBXJ.comXmmc.ToolTipText = txtXMMC.ToolTipText
frmWBXJ.cmdSave.Enabled = True
frmWBXJ.frmTime.Visible = False
frmWBXJ.frmNb.Visible = False
frmWBXJ.cmdD.Visible = False
frmWBXJ.cmdTk.Visible = False
frmWBXJ.cmdCg.Visible = False

End Sub


Private Sub cmdW3_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If Val(cmdW3.ToolTipText) > 0 Then
mod1.BTZ = 36
Call modBJD.BJDWBQing
Call modBJD.BJDBound(cmdW3.ToolTipText, "���̷ְ�")
frmWBXJ.Show
frmWBXJ.cmdSave.Enabled = True
frmWBXJ.frmTime.Visible = False
frmWBXJ.frmNb.Visible = False
frmWBXJ.cmdD.Visible = False
frmWBXJ.cmdTk.Visible = False
frmWBXJ.cmdCg.Visible = False
Exit Sub
End If
ii = MsgBox("�Ƿ��½����̷ְ�ѯ�۵�?", vbInformation + vbYesNo, "Hello!")
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
Set mod1.cmd = New ADODB.command
mod1.cmd.ActiveConnection = mod1.workKK
mod1.cmd.CommandText = "xunJiaAddHT"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@ywy") = mod1.DName
mod1.cmd.Parameters("@uid") = mod1.DHid
mod1.cmd.Parameters("@Lx") = 1
mod1.cmd.Parameters("@zl") = "���̷ְ�"
mod1.cmd.Parameters("@Lcou") = 4 '��������
mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
mod1.cmd.Parameters("@lcRen") = mod1.DName
mod1.cmd.Parameters("@lcUid") = mod1.DHid
mod1.cmd.Parameters("@nLb") = 44
mod1.cmd.Parameters("@xmmc") = txtXMMC.Text
mod1.cmd.Parameters("@xid") = txtXMMC.ToolTipText
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
frmWBXJ.lblBm.Caption = mod1.BM
frmWBXJ.lblQy.Caption = mod1.Qy
frmWBXJ.lblZl.Caption = "���̷ְ�"
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
Set mod1.mA = New ADODB.Recordset
mod1.mA.Close
mod1.mA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.dtgA.DataSource = mod1.mA


'���º�ͬ
tt = "update htping set bid3=" & Val(frmWBXJ.lblBid.Caption) & "where hid=" & Val(lblMHid.Caption)
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
cmdW3.ToolTipText = frmWBXJ.lblBid.Caption


frmWBXJ.cmdBjd.Visible = False
frmWBXJ.txtHg.Locked = True
frmWBXJ.txtYhg.Locked = True
frmWBXJ.txtClf.Locked = True
frmWBXJ.cmdCg.Enabled = False
'frmWBXJ.cmdCong.Visible = False
frmWBXJ.cmdTk.Visible = True
frmWBXJ.Visible = True
frmWBXJ.comXmmc.Text = txtXMMC.Text
frmWBXJ.comXmmc.ToolTipText = txtXMMC.ToolTipText
frmWBXJ.cmdSave.Enabled = True
frmWBXJ.frmTime.Visible = False
frmWBXJ.frmNb.Visible = False
frmWBXJ.cmdD.Visible = False
frmWBXJ.cmdTk.Visible = False
frmWBXJ.cmdCg.Visible = False
End Sub


Private Sub cmdW5_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If Val(cmdW5.ToolTipText) = 0 And txtYwy.ToolTipText = mod1.DHid Then
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
    
'    Set mod1.cmd = New ADODB.command
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
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
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
    mod1.cmd.Parameters("@mt2") = txtXMMC.Text
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
    mod1.cmd.Parameters("@mm2") = txtXMMC.ToolTipText '��Ŀ���
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
    
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        If timZm = 3 Then '����
            cmdW5.Enabled = False
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

    cmdW5.Enabled = False
Set mod1.cmd = Nothing
    

Else '�����ѯ�۵�
    Call modBJD.BJDGXQing
    Call modBJD.BJDBound(Val(cmdW5.ToolTipText), "���")
    Call modBJD.gxbjLocked

    mod1.BTZ = 36
    frmWait.Visible = False
    frmGXBj.Visible = True
    frmGXBj.ZOrder 0
    frmGXBj.cmdMod.Enabled = True
    frmGXBj.cmdSave.Enabled = False
End If
End Sub


Private Sub cmdW6_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
If Val(cmdW6.ToolTipText) = 0 And txtYwy.ToolTipText = mod1.DHid Then
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
    
'    Set mod1.cmd = New ADODB.command
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
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
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
    mod1.cmd.Parameters("@mt2") = txtXMMC.Text
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
    mod1.cmd.Parameters("@mm2") = txtXMMC.ToolTipText '��Ŀ���
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
    Call modBJD.BJDGXQing
    Call modBJD.BJDBound(Val(cmdW6.ToolTipText), "��Ʒ")
    Call modBJD.gxbjLocked

    mod1.BTZ = 36
    frmWait.Visible = False
    frmGXBj.Visible = True
    frmGXBj.ZOrder 0
    frmGXBj.cmdMod.Enabled = True
    frmGXBj.cmdSave.Enabled = False
End If
End Sub

Private Sub cmdYadd_Click()
Dim tt As String
Dim hg As Single
On Error Resume Next
If Val(txtFED.Text) = 0 Or Val(txtYingFu.Text) = 0 Then
Exit Sub
End If

tt = "select yjff from htping where htbh='" & txtHtbh.Text & "'"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.HTP.RecordCount) Or mod1.HTP.RecordCount = 0 Then
    Exit Sub
End If
If mod1.HTP.Fields("yjff").Value = True Then
    MsgBox ("�����Ѿ�ȫ��֧��,�����ٸ���!")
    Exit Sub
End If

Set mod1.cmd = New ADODB.command
mod1.cmd.ActiveConnection = mod1.CC
mod1.cmd.CommandText = "htyjAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@htbh") = Trim(txtHtbh.Text)
mod1.cmd.Parameters("@YED") = Val(txtFED.Text) / 100
mod1.cmd.Parameters("@yingFu") = Val(txtYingFu.Text)
mod1.cmd.Parameters("@xmmc") = Trim(txtXMMC.Text)
mod1.cmd.Execute
Set cmd = Nothing
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
'HG = HG + Val(txtYingFu.Text)
'If HG > Val(txtYj.Text) Then
'    MsgBox "��д�������!"
'    txtYingFu.Text = ""
'    Exit Sub
'End If
'End If
txtYj1.Text = hg
txtLr1.Text = Val(txtJlr1.Text) - Val(txtYj1.Text)
tt = "update htping set yj=" & Val(txtYj1.Text) & ",xmlr=" & Val(txtLr1.Text) & " where htbh='" & txtHtbh.Text & "'"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
End Sub

Private Sub cmdYdel_Click()
Dim tt As String
Dim hg As Single
Dim ii As Integer
Dim Yid As Long
Dim Ywy As String
On Error Resume Next
MMdtgYJ.Col = 4
Ywy = MMdtgYJ.Text
MMdtgYJ.Col = 3
Yid = 0
Yid = MMdtgYJ.Text


If Yid = 0 Then
Exit Sub
End If

If Ywy <> "" Then
    MsgBox "�˵��Ѿ�����,����ɾ��! ���ȷ��Ҫɾ��,������������ϵ!"
    Exit Sub
End If


ii = MsgBox("�Ƿ�ɾ���˼�¼?", vbQuestion + vbYesNo, "ѯ��")
If ii = vbNo Then
    Exit Sub
End If
tt = "delete from yongjin where yid=" & Yid
Set mod1.HTP = New ADODB.Recordset
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
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
End Sub


Private Sub Command3_Click()

End Sub

Private Sub dt3_CloseUp()
txtF.Text = dt3.Value
End Sub


Private Sub dt4_CloseUp()
txtL.Text = dt4.Value
End Sub


Private Sub mmdtgbao_Click()
Dim tt As String
Dim liD As Long
On Error Resume Next
MMdtgBao.Col = 11
txtTl.Text = MMdtgBao.Text
MMdtgBao.Col = 12
txtDj.Text = MMdtgBao.Text
MMdtgBao.Col = 16
liD = MMdtgBao.Text
tt = "select * from xunJiaMxView where lid=" & liD
mod1.mGx.Close
mod1.mGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set MMdtgMa.DataSource = mod1.mGx
End Sub

Private Sub mmdtgbao_RowColChange()
Dim tt As String
Dim liD As Long
On Error Resume Next
MMdtgBao.Col = 11
txtTl.Text = MMdtgBao.Text
MMdtgBao.Col = 12
txtDj.Text = MMdtgBao.Text
MMdtgBao.Col = 16
liD = MMdtgBao.Text
tt = "select * from xunJiaMxView where lid=" & liD
mod1.mGx.Close
mod1.mGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set MMdtgMa.DataSource = mod1.mGx
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
MMdtgFk.Col = 2
txtYed.Text = Val(MMdtgFk.Text) * 100
MMdtgFk.Col = 5
lblFid.Caption = MMdtgFk.Text
End Sub

Private Sub mmdtgfk_RowColChange()
On Error Resume Next
If Val(MMdtgFk.Text) = 0 Then Exit Sub
MMdtgFk.Col = 1
txtYrq.Text = MMdtgFk.Text
MMdtgFk.Col = 2
txtYed.Text = Val(MMdtgFk.Text) * 100
MMdtgFk.Col = 5
lblFid.Caption = MMdtgFk.Text
End Sub


Private Sub dtpYf_CloseUp()
txtYrq.Text = dtpYf.Value
End Sub

Private Sub Form_Click()
frmQm.Visible = False
lblTX.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 76 Then
    If mod1.Kyj = True Then
        If frmYj.Visible = False Then
            frmYj.Visible = True
            lblTcBe.Visible = True
            txtTcBe.Visible = True
        Else
            frmYj.Visible = False
            lblTcBe.Visible = False
            txtTcBe.Visible = False
        End If
   End If
    
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
MsgBox ("��")
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
Me.Left = 0
Me.Top = 0
frmJi.BorderStyle = 0

'''''''''Set mWb = New ADODB.Recordset
'''''''''Set mLj = New ADODB.Recordset
''''''''''Set adoOid = New ADODB.Recordset
'''''''''Set mod1.mBx = New ADODB.Recordset
'''''''''Set mod1.mGx = New ADODB.Recordset
'''''''''Set mod1.mFk = New ADODB.Recordset
'''''''''Set mod1.mYj = New ADODB.Recordset
'''''''''Set mod1.mBxCP = New ADODB.Recordset
'''''''''Set mod1.mGxCP = New ADODB.Recordset
'''''''''
'''''''''Set mod1.mA = New ADODB.Recordset
'''''''''Set mod1.mB = New ADODB.Recordset

MMdtgMa.ColWidth(0) = 300
''MMdtgMa.ColWidth(8) = 2000
''MMdtgMa.ColWidth(15) = 0
''MMdtgMa.ColWidth(16) = 0
MMdtgBao.ColWidth(0) = 300
'''MMdtgBao.ColWidth(8) = 2000
'''MMdtgBao.ColWidth(15) = 0
'''MMdtgBao.ColWidth(16) = 0
MMdtgBao.Left = 0
MMdtgBao.Top = 0
frmYj.BorderStyle = 0


MMdtgA.ColWidth(0) = 300
MMdtgA.ColWidth(2) = 2000
MMdtgA.ColWidth(3) = 700
MMdtgA.ColWidth(4) = 0

MMdtgFk.ColWidth(0) = 300
MMdtgFk.ColWidth(4) = 0
MMdtgFk.ColWidth(5) = 0
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
If htBrow.Visible = True Then
    htBrow.Enabled = True
    htBrow.ZOrder 0
ElseIf Dialog.Enabled = True Then
    Dialog.ZOrder 0
    Dialog.Enabled = True
End If
Cancel = True
End Sub

Private Sub tabGc_Click(PreviousTab As Integer)
Dim oo As Integer
For oo = 0 To 5
frmgc(oo).Visible = False
Next
frmgc(tabGc.Tab).Visible = True
End Sub

Private Sub tabHt_Click(PreviousTab As Integer)
frmQm.Visible = False

End Sub

Private Sub tabHt_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 76 Then
    'If mod1.Kyj = True Then
        If frmYj.Visible = False Then
            frmYj.Visible = True
            lblTcBe.Visible = True
            txtTcBe.Visible = True
        Else
            frmYj.Visible = False
            lblTcBe.Visible = False
            txtTcBe.Visible = False
        End If
   ' End If
    
End If
End Sub


Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 2 Then '���Ϊ��Ӻ�ͬ����
    Call modNewHT.NewLocked
    cmdSave.Enabled = False
    If Val(lblLc.Caption) = 0 Then
        lblLc.Caption = 1
    End If
ElseIf timZm = 10 Then 'ǩ��
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
ElseIf timZm = 11 Then
    cmdHT.Visible = False
    
End If
timQuit.Enabled = False

End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2 from ml where zid=" & mod1.Zid
Set mod1.WP = New ADODB.Recordset
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '�ύ�ɹ�
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 2 Then
        cmdSave.Enabled = False
    ElseIf timZm = 3 Then
        
        frmGXBj.lblBid.Caption = mod1.WP.Fields("mm1").Value
        frmGXBj.lblBh.Caption = "XJD" & mod1.WP.Fields("mm1").Value
        frmGXBj.lblLcou.Caption = 3 '��������
        frmGXBj.lblLc.Caption = 0
        frmGXBj.lblLcRen.Caption = mod1.DName
        frmGXBj.lblLcUid.Caption = mod1.DHid
        frmGXBj.lblNlb.Caption = 43
        frmGXBj.lblYwy.Caption = mod1.DName
        frmGXBj.lblUid.Caption = mod1.DHid
        frmGXBj.lblZl.Caption = mod1.WP.Fields("mt1").Value
        If mod1.WP.Fields("mt1").Value = "���" Then
            cmdW5.ToolTipText = mod1.WP.Fields("mm1").Value
        ElseIf mod1.WP.Fields("mt1").Value = "��Ʒ" Then
            cmdW6.ToolTipText = mod1.WP.Fields("mm1").Value
        End If
        frmGXBj.comXmmc.Text = txtXMMC.Text
        frmGXBj.comXmmc.ToolTipText = txtXMMC.ToolTipText
        frmGXBj.txtHg.Locked = True
        frmGXBj.txtYhg.Locked = True
        frmGXBj.lblHtbh.Caption = FMXC.lblMHid.Caption
        
            '�������̰�ť
            Call modBJD.XJGXLcNew(43)
            

        frmGXBj.cmdMod.Enabled = False
        frmGXBj.frmCg.Enabled = False
        'ˢ�¹����б�
        tt = "select * from xunJIamxView where bid=" & Val(frmGXBj.lblBid.Caption)
            mod1.mGx.Close
            mod1.mGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            If IsNull(mod1.mGx.RecordCount) = True Then
                MsgBox ("��ȡ��������,���ڹرպ�����һ��!")
            End If
            Set frmGXBj.dtgMa.DataSource = mod1.mGx
        
        frmGXBj.cmdSave.Enabled = True
        frmGxBiao.Enabled = False
        'frmGXBj.cmdBjd.Visible = False
        frmGXBj.txtYhg.Locked = True
        frmGXBj.comXmmc.Locked = False
        frmGXBj.lblZl.ForeColor = &HC000C0
        frmGXBj.lblzlZ.ForeColor = &HC000C0
        frmGXBj.txtMj.Locked = True
        frmGXBj.txtDj.Locked = True
        
        mod1.BTZ = 36
        frmGXBj.Visible = True
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
    
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then 'δ���

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '����ʧ��
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

Private Sub txtW1_DblClick()
frmWai.Visible = True
End Sub


Private Sub txtYj1_Click()
frmYM.Visible = True
End Sub
