VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form wbHTP 
   Caption         =   "ά����ά�޺�ͬ����"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   9225
   Begin VB.CommandButton chkD 
      Height          =   315
      Left            =   3900
      TabIndex        =   109
      Top             =   8460
      Width           =   855
   End
   Begin VB.CommandButton chkB 
      Height          =   315
      Left            =   2970
      TabIndex        =   108
      Top             =   8460
      Width           =   915
   End
   Begin VB.CommandButton chkC 
      Height          =   315
      Left            =   2040
      TabIndex        =   107
      Top             =   8460
      Width           =   855
   End
   Begin VB.CommandButton chkE 
      Height          =   315
      Left            =   1050
      TabIndex        =   106
      Top             =   8460
      Width           =   915
   End
   Begin VB.CommandButton chkA 
      Height          =   315
      Left            =   60
      TabIndex        =   105
      Top             =   8460
      Width           =   915
   End
   Begin VB.TextBox txtTcRQ 
      Height          =   315
      Left            =   7110
      Locked          =   -1  'True
      TabIndex        =   104
      Text            =   "���ȡ������"
      Top             =   7110
      Width           =   1845
   End
   Begin VB.CommandButton cmdCount 
      Caption         =   "����"
      Height          =   315
      Left            =   6240
      TabIndex        =   103
      Top             =   6750
      Width           =   705
   End
   Begin VB.TextBox txtTcBe 
      Height          =   285
      Left            =   5640
      TabIndex        =   100
      Text            =   "6"
      Top             =   6750
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox txtXMNr 
      Height          =   2895
      Left            =   7140
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   99
      Top             =   3720
      Width           =   1845
   End
   Begin VB.ComboBox txtKhmc 
      Height          =   300
      Left            =   1320
      TabIndex        =   90
      Text            =   "txtKhmc"
      ToolTipText     =   "�����б���ѡ��ͻ�"
      Top             =   570
      Width           =   3345
   End
   Begin VB.TextBox txtXMMC 
      Height          =   285
      Left            =   5850
      TabIndex        =   89
      Top             =   540
      Width           =   3105
   End
   Begin VB.TextBox txtKhdm 
      Height          =   270
      Left            =   1350
      TabIndex        =   87
      Top             =   1020
      Width           =   1365
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
      Height          =   270
      Left            =   1290
      TabIndex        =   84
      Top             =   5820
      Width           =   3105
   End
   Begin VB.TextBox txtJlr2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   83
      Top             =   5820
      Width           =   2235
   End
   Begin VB.CommandButton cmdFkqk 
      Caption         =   "�������"
      Height          =   285
      Left            =   4710
      TabIndex        =   81
      Top             =   2790
      Width           =   4275
   End
   Begin VB.TextBox txtADR 
      Height          =   285
      Left            =   5850
      TabIndex        =   80
      Top             =   2340
      Width           =   3105
   End
   Begin VB.Frame frmHtxz 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   255
      Left            =   1350
      TabIndex        =   75
      Top             =   1695
      Width           =   3045
      Begin VB.OptionButton optA 
         Caption         =   "C. ά����ͬ"
         Height          =   225
         Index           =   3
         Left            =   0
         TabIndex        =   77
         Top             =   30
         Width           =   1305
      End
      Begin VB.OptionButton optA 
         Caption         =   "D. ά�޺�ͬ"
         Height          =   255
         Index           =   4
         Left            =   1470
         TabIndex        =   76
         Top             =   0
         Width           =   1305
      End
   End
   Begin VB.TextBox txtCbze2 
      Height          =   285
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   3720
      Width           =   2235
   End
   Begin VB.TextBox txtCbze1 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1290
      TabIndex        =   72
      Top             =   3720
      Width           =   3105
   End
   Begin VB.TextBox txtGLG 
      Height          =   285
      Left            =   5850
      TabIndex        =   70
      Top             =   1980
      Width           =   3105
   End
   Begin VB.Frame Frame1 
      Caption         =   "��Ʊ���ͣ�"
      Height          =   765
      Left            =   360
      TabIndex        =   66
      Top             =   6690
      Width           =   4035
      Begin VB.CommandButton cmdJi 
         Caption         =   "����"
         Height          =   495
         Left            =   3570
         TabIndex        =   78
         Top             =   180
         Width           =   345
      End
      Begin VB.OptionButton optLc 
         Caption         =   "����Ʊ"
         Height          =   195
         Left            =   2370
         TabIndex        =   69
         Top             =   300
         Width           =   1065
      End
      Begin VB.OptionButton optLb 
         Caption         =   "��ҵ��Ʊ"
         Height          =   195
         Left            =   1260
         TabIndex        =   68
         Top             =   300
         Width           =   1065
      End
      Begin VB.OptionButton optLa 
         Caption         =   "��ֵ��Ʊ"
         Height          =   195
         Left            =   180
         TabIndex        =   67
         Top             =   300
         Width           =   1065
      End
   End
   Begin VB.TextBox txtJy 
      Height          =   555
      Left            =   1290
      TabIndex        =   65
      Top             =   7530
      Width           =   7395
   End
   Begin VB.TextBox txtTc2 
      Height          =   285
      Left            =   5640
      TabIndex        =   63
      Top             =   7110
      Width           =   1305
   End
   Begin VB.TextBox txtLr2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   6360
      Width           =   2235
   End
   Begin VB.TextBox txtQt2 
      Height          =   270
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   61
      ToolTipText     =   "˫���˴����Կ���Ŀ�����嵥"
      Top             =   5550
      Width           =   2235
   End
   Begin VB.TextBox txtYj2 
      Height          =   270
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   6090
      Width           =   2235
   End
   Begin VB.TextBox txtFbje2 
      Height          =   285
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   4980
      Width           =   2235
   End
   Begin VB.TextBox txtYf2 
      Height          =   270
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   5280
      Width           =   2235
   End
   Begin VB.TextBox txtCLF2 
      Height          =   315
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   4650
      Width           =   2235
   End
   Begin VB.TextBox txtRgf2 
      Height          =   315
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   4320
      Width           =   2235
   End
   Begin VB.TextBox txtClcb2 
      Height          =   285
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   55
      ToolTipText     =   "˫���˴����Կ����ϳɱ��嵥"
      Top             =   4020
      Width           =   2235
   End
   Begin VB.TextBox txtClcb1 
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   4020
      Width           =   3105
   End
   Begin VB.TextBox txtQt1 
      Height          =   270
      Left            =   1290
      TabIndex        =   48
      Top             =   5550
      Width           =   3105
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
      Height          =   270
      Left            =   1290
      TabIndex        =   45
      Top             =   6360
      Width           =   3105
   End
   Begin VB.TextBox txtYf1 
      Height          =   270
      Left            =   1290
      TabIndex        =   42
      Top             =   5280
      Width           =   3105
   End
   Begin VB.TextBox txtYj1 
      Height          =   270
      Left            =   1290
      TabIndex        =   41
      Top             =   6090
      Width           =   3105
   End
   Begin VB.TextBox txtFbje1 
      Height          =   285
      Left            =   1290
      TabIndex        =   39
      Top             =   4980
      Width           =   3105
   End
   Begin VB.TextBox txtCLF1 
      Height          =   315
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   4650
      Width           =   3105
   End
   Begin VB.TextBox txtRgf1 
      Height          =   315
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   4320
      Width           =   3105
   End
   Begin VB.Frame frmZt 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   885
      Left            =   4800
      TabIndex        =   27
      Top             =   8310
      Width           =   1185
      Begin VB.OptionButton optG 
         Caption         =   "�� �� ��"
         Height          =   195
         Left            =   90
         TabIndex        =   91
         Top             =   270
         Width           =   1035
      End
      Begin VB.OptionButton optP 
         Caption         =   "����׶�"
         Height          =   180
         Left            =   90
         TabIndex        =   30
         Top             =   60
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optZ 
         Caption         =   "ִ�н׶�"
         Height          =   225
         Left            =   90
         TabIndex        =   29
         Top             =   480
         Width           =   1035
      End
      Begin VB.OptionButton optW 
         Caption         =   "ִ�����"
         Height          =   225
         Left            =   90
         TabIndex        =   28
         Top             =   690
         Width           =   1035
      End
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
      Left            =   1320
      TabIndex        =   25
      Top             =   2790
      Width           =   3105
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ��"
      Enabled         =   0   'False
      Height          =   585
      Left            =   7950
      Picture         =   "wbHTPOld.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8610
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "�޸�"
      Height          =   585
      Left            =   6600
      Picture         =   "wbHTPOld.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8610
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   585
      Left            =   8610
      Picture         =   "wbHTPOld.frx":05CC
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8610
      Width           =   585
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "�ύ"
      Height          =   585
      Left            =   7260
      Picture         =   "wbHTPOld.frx":06CE
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8610
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ӡ"
      Height          =   585
      Left            =   5940
      Picture         =   "wbHTPOld.frx":0D38
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8610
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtMon 
      Height          =   270
      Left            =   1320
      TabIndex        =   17
      Top             =   2400
      Width           =   945
   End
   Begin VB.CommandButton cmdWb 
      Caption         =   "�ͻ�����"
      Height          =   315
      Left            =   2880
      TabIndex        =   11
      Top             =   990
      Width           =   1545
   End
   Begin VB.TextBox txtHtdate 
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dddddd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   210
      Left            =   5880
      TabIndex        =   4
      Top             =   1305
      Width           =   2655
   End
   Begin VB.TextBox txtHtbh 
      Height          =   270
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1350
      Width           =   3105
   End
   Begin VB.TextBox txtYwy 
      Height          =   315
      Left            =   5850
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   870
      Width           =   3105
   End
   Begin VB.ComboBox comQy 
      Height          =   300
      ItemData        =   "wbHTPOld.frx":13A2
      Left            =   5850
      List            =   "wbHTPOld.frx":13A4
      TabIndex        =   1
      Text            =   "comQy"
      Top             =   1590
      Width           =   3105
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   255
      Left            =   5850
      TabIndex        =   12
      Top             =   1260
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "yyyy��M��d��"
      Format          =   149880835
      CurrentDate     =   38098.7575810185
   End
   Begin MSComCtl2.DTPicker dt4 
      Height          =   315
      Left            =   3090
      TabIndex        =   15
      Top             =   2040
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      Format          =   149880833
      CurrentDate     =   38098
   End
   Begin MSComCtl2.DTPicker dt3 
      Height          =   315
      Left            =   1320
      TabIndex        =   16
      Top             =   2040
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      Format          =   149880833
      CurrentDate     =   38098
   End
   Begin MSComCtl2.UpDown UpDa 
      Height          =   315
      Left            =   5970
      TabIndex        =   102
      Top             =   6750
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   503
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label lblHid 
      Height          =   315
      Left            =   7020
      TabIndex        =   110
      Top             =   8190
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblTcBe 
      Caption         =   "��ɱ���"
      Height          =   195
      Left            =   4710
      TabIndex        =   101
      Top             =   6810
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label23 
      Caption         =   "��Ŀ����"
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
      Left            =   7380
      TabIndex        =   98
      Top             =   3420
      Width           =   945
   End
   Begin VB.Label lblZj 
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   3840
      TabIndex        =   97
      Top             =   8790
      Width           =   885
   End
   Begin VB.Label lblYz 
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   1965
      TabIndex        =   96
      Top             =   8790
      Width           =   885
   End
   Begin VB.Label lblJl 
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   2895
      TabIndex        =   95
      Top             =   8790
      Width           =   885
   End
   Begin VB.Label lblYw 
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   30
      TabIndex        =   94
      Top             =   8790
      Width           =   885
   End
   Begin VB.Label lblJz 
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   990
      TabIndex        =   93
      Top             =   8790
      Width           =   885
   End
   Begin VB.Label lblBM 
      Caption         =   "Label27"
      Height          =   315
      Left            =   0
      TabIndex        =   92
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "��Ŀ����"
      Height          =   225
      Left            =   4710
      TabIndex        =   88
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label6 
      Caption         =   "�ͻ�����"
      Height          =   255
      Left            =   330
      TabIndex        =   86
      Top             =   1050
      Width           =   885
   End
   Begin VB.Label lblJlr 
      Caption         =   "�� �� 1"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   360
      TabIndex        =   85
      Top             =   5880
      Width           =   915
   End
   Begin VB.Label Label28 
      Caption         =   "����֧��"
      Height          =   255
      Left            =   1140
      TabIndex        =   82
      Top             =   8220
      Width           =   765
   End
   Begin VB.Label Label26 
      Caption         =   "��Ŀ��ַ"
      Height          =   255
      Left            =   4710
      TabIndex        =   79
      Top             =   2370
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
      Left            =   360
      TabIndex        =   73
      Top             =   3780
      Width           =   885
   End
   Begin VB.Label Label16 
      Caption         =   "��ҵ����"
      Height          =   255
      Left            =   4710
      TabIndex        =   71
      Top             =   2010
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "������"
      Height          =   285
      Left            =   360
      TabIndex        =   64
      Top             =   7620
      Width           =   765
   End
   Begin VB.Label Label17 
      Caption         =   "���ϳɱ�"
      Height          =   255
      Left            =   360
      TabIndex        =   54
      Top             =   4050
      Width           =   825
   End
   Begin VB.Label Label11 
      Caption         =   "ʵ ��"
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
      Left            =   4740
      TabIndex        =   52
      Top             =   3360
      Width           =   675
   End
   Begin VB.Label Label9 
      Caption         =   "Ԥ ��"
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
      Left            =   1890
      TabIndex        =   51
      Top             =   3360
      Width           =   675
   End
   Begin VB.Label Label8 
      Caption         =   "�� Ŀ"
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
      Left            =   360
      TabIndex        =   50
      Top             =   3360
      Width           =   675
   End
   Begin VB.Label Label20 
      Caption         =   "��Ŀ����"
      Height          =   225
      Left            =   360
      TabIndex        =   49
      Top             =   5610
      Width           =   885
   End
   Begin VB.Label lblLr 
      Caption         =   "�� �� 2"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   360
      TabIndex        =   47
      Top             =   6420
      Width           =   915
   End
   Begin VB.Label lblTC 
      Caption         =   "��    ��"
      Height          =   195
      Left            =   4710
      TabIndex        =   46
      Top             =   7170
      Width           =   735
   End
   Begin VB.Label Label19 
      Caption         =   "��    ��"
      Height          =   195
      Left            =   360
      TabIndex        =   44
      Top             =   5340
      Width           =   855
   End
   Begin VB.Label lblYj 
      Caption         =   "��    ��"
      Height          =   225
      Left            =   360
      TabIndex        =   43
      Top             =   6150
      Width           =   975
   End
   Begin VB.Label Label18 
      Caption         =   "�ְ����"
      Height          =   195
      Left            =   360
      TabIndex        =   40
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "�� �� ��"
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   4710
      Width           =   915
   End
   Begin VB.Label Label15 
      Caption         =   "�� �� ��"
      Height          =   285
      Left            =   360
      TabIndex        =   35
      Top             =   4380
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "ҵ��Ա"
      Height          =   225
      Left            =   210
      TabIndex        =   34
      Top             =   8220
      Width           =   615
   End
   Begin VB.Label Label35 
      Caption         =   "���۾���"
      Height          =   255
      Left            =   3015
      TabIndex        =   33
      Top             =   8220
      Width           =   765
   End
   Begin VB.Label Label36 
      Caption         =   "������"
      Height          =   255
      Left            =   2025
      TabIndex        =   32
      Top             =   8220
      Width           =   855
   End
   Begin VB.Label Label37 
      Caption         =   "�ܾ���"
      Height          =   255
      Left            =   3930
      TabIndex        =   31
      Top             =   8220
      Width           =   645
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
      Left            =   120
      TabIndex        =   26
      Top             =   2850
      Width           =   1095
   End
   Begin VB.Label Label12 
      Caption         =   "��"
      Height          =   255
      Left            =   2370
      TabIndex        =   19
      Top             =   2430
      Width           =   195
   End
   Begin VB.Label Label10 
      Caption         =   "ά�ޱ�����"
      Height          =   225
      Left            =   150
      TabIndex        =   18
      Top             =   2430
      Width           =   1065
   End
   Begin VB.Label Label21 
      Caption         =   "---��"
      Height          =   225
      Left            =   2700
      TabIndex        =   14
      Top             =   2100
      Width           =   375
   End
   Begin VB.Label Label27 
      Caption         =   "ά�޹���"
      Height          =   225
      Left            =   330
      TabIndex        =   13
      Top             =   2085
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "��ͬ����"
      Height          =   195
      Left            =   330
      TabIndex        =   10
      Top             =   1755
      Width           =   915
   End
   Begin VB.Label Label25 
      Caption         =   "��ͬ���"
      Height          =   225
      Left            =   330
      TabIndex        =   9
      Top             =   1410
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "��    ��"
      Height          =   255
      Left            =   4710
      TabIndex        =   8
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "ҵ �� Ա"
      Height          =   255
      Left            =   4710
      TabIndex        =   7
      Top             =   930
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "�ͻ�����"
      Height          =   225
      Left            =   330
      TabIndex        =   6
      Top             =   630
      Width           =   855
   End
   Begin VB.Label Label44 
      Caption         =   "��    ��"
      Height          =   255
      Left            =   4710
      TabIndex        =   5
      Top             =   1650
      Width           =   855
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "ά����ά�޺�ͬ����"
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
      Left            =   2850
      TabIndex        =   0
      Top             =   60
      Width           =   3405
   End
End
Attribute VB_Name = "wbHTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tmpHtze As Single '��ʱ��ͬ�ܶ�

Private Sub chkC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If txtKhmc.Text = "" Then
MsgBox "����ȷ����ͻ�����"
txtKhmc.SetFocus
Exit Sub
End If
If comQy.Text = "" Then
MsgBox "����ȷѡ������"
comQy.SetFocus
Exit Sub
End If
If txtGLG.Text = "" Then
MsgBox "����ȷ�������˾"
txtGLG.SetFocus
Exit Sub
End If
If txtMon.Text = "" And optA(4).Value = True Then
MsgBox "����ȷ����ά�ޱ�����"
txtMon.SetFocus
Exit Sub
End If
If txtCbze1.Text = "" Then
MsgBox "����ȷ����ɱ��ܶ�"
Exit Sub
End If
If txtRgf1.Text = "" Then
MsgBox "����ȷ�����˹���"
Exit Sub
End If
If txtCLF1.Text = "" Then
MsgBox "����ȷ������÷�"
Exit Sub
End If
'If txtYj1.Text = "" Then
'MsgBox "����ȷ��дӶ��"
''txtYj1.SetFocus
'Exit Sub
'End If
If txtQt1.Text = "" Then
MsgBox "����ȷ������Ŀ����"
Exit Sub
End If
If txtLr1.Text = "" Then
MsgBox "����ȷ��дë��"
txtLr1.SetFocus
Exit Sub
End If

If frmFuK.adoHpt.Recordset.RecordCount = 0 Then
MsgBox "����ȷ��дӦ�տ�"
Exit Sub
End If


End Sub


Private Sub cmdBack_Click()
'On Error Resume Next
'Dim tt As String
''khAdd.Close
'Dim ii As Integer
'
Call mod1.DelDKZ  '�˳���ʱɾ���򿪼�¼,���ñ����ܴ򿪴˵���
'

If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf htBrow.Visible = True Then
    htBrow.Enabled = True
    htBrow.ZOrder 0
ElseIf htBrowG.Visible = True Then
    htBrowG.Enabled = True
    htBrow.ZOrder 0
End If
'htBrow.MousePointer = 0
wbHTP.Visible = False
End Sub

Private Sub cmdFkQ_Click()
frmFuK.Show
End Sub

Private Sub cmdFkqk_Click()
If txtJlr1.Text = "" Then
MsgBox "���ȼ��������(�ں�ͬ�ܽ����س���)"
txtHtze.SetFocus
Exit Sub
End If
If Val(txtHtze.Text) > 0 Then
wbMx.Show
wbMx.SSTab1.Tab = 3
wbMx.SSTab1.Enabled = True
End If
wbMx.lblHtze.Caption = txtHtze.Text
End Sub

Private Sub cmdJi_Click()
On Error Resume Next
'����ɱ��ܶ�
txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
Val(txtYf1.Text) + Val(txtQt1.Text)

'��������
    If optLa.Value = True Or optLb.Value = True Then
        txtJlr1.Text = Round(Val(txtHtze.Text) / 1.17 - Val(txtCbze1.Text), 2)
        txtLr1.Text = Round(Val(txtHtze.Text) / 1.17 - Val(txtCbze1.Text) - Val(txtYj1.Text), 2)
    ElseIf optLc.Value = True Then
        txtJlr1.Text = Round(Val(txtHtze.Text) / 1.06 - Val(txtCbze1.Text), 2)
        txtLr1.Text = Round(Val(txtHtze.Text) / 1.06 - Val(txtCbze1.Text) - Val(txtYj1.Text), 2)
    End If
wbMx.lblHtze.Caption = txtHtze.Text
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub dt1_CloseUp()
txtHtdate.Text = dt1.Value
End Sub

Private Sub dtgKhmc_DblClick()


End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 76 Then
    wbHTP.txtYj1.Locked = True
    wbHTP.txtYj2.Locked = True
 If wbHTP.lblYj.Visible = False Then
'        If (mod1.Kyj = True And mod1.DName = chkB.Tag) Or _
'        (mod1.DName = "�Ŵ���" And (optW.Value = True Or optZ.Visible = True)) Or _
'         mod1.ZW = "�ܾ���" Or mod1.ZW = "���ܾ���" Or (mod1.Kyj = True And mod1.BMN = wbHTP.lblBM.Caption And mod1.Qy = wbHTP.comQy.Text) Then
    If mod1.Kyj = True Then
           'Ӷ������2�������ʾ
            wbHTP.txtYj1.Visible = True
            wbHTP.txtYj2.Visible = True
            wbHTP.txtLr1.Visible = True
            wbHTP.txtLr2.Visible = True
            wbHTP.txtTc2.Visible = True
            wbHTP.lblYj.Visible = True
            wbHTP.lblLr.Visible = True
            wbHTP.lblTC.Visible = True
            wbHTP.lblTcBe.Visible = True
            wbHTP.txtTcBe.Visible = True
            wbHTP.UpDa.Visible = True
'            If mod1.KY2 = True And optW.Value = False Then 'С��ֻ���޸ĺ�ͬĩ��ɵ�ʵ��Ӷ��
'                txtYj2.Locked = False
'            End If
'            If mod1.KY1 = True Then '���۾������ϰ�ǩ�ֺ�,�Ͳ����޸�Ԥ��Ӷ��
'                If chkD.Caption = "" Or mod1.ZW = "�ܾ���" Or mod1.ZW = "���ܾ���" Then
'                    txtYj1.Locked = False
'                End If
'            End If
    End If


  Else
        'Ӷ������2�������ʾ
    wbHTP.txtYj1.Visible = False
    wbHTP.txtYj2.Visible = False
    wbHTP.txtLr1.Visible = False
    wbHTP.txtLr2.Visible = False
    wbHTP.lblTcBe.Visible = False
    wbHTP.txtTcBe.Visible = False
    wbHTP.UpDa.Visible = False
    wbHTP.lblYj.Visible = False
    wbHTP.lblLr.Visible = False
    wbHTP.lblTC.Visible = False
  End If
    
End If
End Sub

Private Sub Form_Load()
Dim tt As String
Dim oo As Integer
wbHTP.Width = 9345
wbHTP.Height = 9630
wbHTP.Top = 0
wbHTP.Left = 3000



''��������
'tt = "Select * from yzQy"
'frmAdo.adoTmp.Recordset.Close
'frmAdo.adoTmp.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic
'For oo = wbHTP.comQy.ListCount - 1 To 0 Step -1
'wbHTP.comQy.RemoveItem oo
'Next
'frmAdo.adoTmp.Recordset.MoveFirst
'For oo = 0 To frmAdo.adoTmp.Recordset.RecordCount - 1
'wbHTP.comQy.AddItem frmAdo.adoTmp.Recordset.Fields("qy").Value, oo
'frmAdo.adoTmp.Recordset.MoveNext
'Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'dtgKhmc.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

If MDI.Cq = False Then
Call mod1.DelDKZ ' '�˳���ʱɾ���򿪼�¼,���ñ����ܴ򿪴˵���
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf htBrow.Visible = True Then
    htBrow.Enabled = True
    htBrow.ZOrder 0
ElseIf htBrowG.Visible = True Then
    htBrowG.Enabled = True
    htBrow.ZOrder 0
End If
wbHTP.Visible = False
End If
End Sub

Private Sub optLa_Click()
'On Error Resume Next
'    cmdJi.Enabled = True
'
'            '����ɱ��ܶ�
''txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
'Val(txtYf1.Text) + Val(txtYj1.Text) + Val(txtQt1.Text)
'txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
'Val(txtYf1.Text) + Val(txtQt1.Text)
'
''��������
'        txtJlr1.Text = Round(Val(txtHtze.Text) / 1.17 - Val(txtCbze1.Text), 2)
'        txtLr1.Text = Round(Val(txtHtze.Text) / 1.17 - Val(txtCbze1.Text) - Val(txtYj1.Text), 2)
'
'wbMx.lblHtZe.Caption = txtHtze.Text
'''����Ӧ�ձ��еĽ��
''frmFuK.adoHpt.Recordset.MoveFirst
''Do While Not frmFuK.adoHpt.Recordset.EOF
''frmFuK.adoHpt.Recordset.Fields("yingfJe").Value = frmFuK.adoHpt.Recordset.Fields("ED") _
''* Val(wbMx.lblHtZe.Caption)
''frmFuK.adoHpt.Recordset.MoveNext
''Loop
End Sub

Private Sub optLb_Click()
'On Error Resume Next
'    cmdJi.Enabled = True
'
'        '����ɱ��ܶ�
''txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
'Val(txtYf1.Text) + Val(txtYj1.Text) + Val(txtQt1.Text)
'txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
'Val(txtYf1.Text) + Val(txtQt1.Text)
'
''��������
'        txtJlr1.Text = Round(Val(txtHtze.Text) / 1.17 - Val(txtCbze1.Text), 2)
'
'        txtLr1.Text = Round(Val(txtHtze.Text) / 1.17 - Val(txtCbze1.Text) - Val(txtYj1.Text), 2)
'
'wbMx.lblHtZe.Caption = txtHtze.Text
'''����Ӧ�ձ��еĽ��
''frmFuK.adoHpt.Recordset.MoveFirst
''Do While Not frmFuK.adoHpt.Recordset.EOF
''frmFuK.adoHpt.Recordset.Fields("yingfJe").Value = frmFuK.adoHpt.Recordset.Fields("ED") _
''* Val(wbMx.lblHtZe.Caption)
''frmFuK.adoHpt.Recordset.MoveNext
''Loop

End Sub

Private Sub optLc_Click()
'On Error Resume Next
'    cmdJi.Enabled = True
'
'    '����ɱ��ܶ�
''txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
'Val(txtYf1.Text) + Val(txtYj1.Text) + Val(txtQt1.Text)
'txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
'Val(txtYf1.Text) + Val(txtQt1.Text)
''��������
'        txtJlr1.Text = Round(Val(txtHtze.Text) / 1.06 - Val(txtCbze1.Text), 2)
'
'        txtLr1.Text = Round(Val(txtHtze.Text) / 1.06 - Val(txtCbze1.Text) - Val(txtYj1.Text), 2)
'
'wbMx.lblHtZe.Caption = txtHtze.Text


End Sub

Private Sub optZ_GotFocus()
cmdSave.Enabled = True
End Sub

Private Sub txtClcb1_DblClick()
If txtKhmc.Text <> "" Then
    wbMx.Show
    wbMx.cmdMod1.Enabled = True
    wbMx.SSTab1.Tab = 2
    If Val(txtHtze.Text) > 0 Then
    wbMx.SSTab1.TabEnabled(3) = True
    'wbMx.cmdMod1.Enabled = False
    Else
    wbMx.SSTab1.TabEnabled(3) = True
    
    End If
End If
wbMx.SSTab1.Enabled = True
End Sub

Private Sub txtClcb2_DblClick()
wbMx.Show

wbMx.SSTab1.Tab = 2
If Val(txtHtze.Text) > 0 Then
wbMx.SSTab1.TabEnabled(3) = True
Else
wbMx.SSTab1.TabEnabled(3) = False
End If
End Sub


Private Sub txtCLF_DblClick()
wbMx.Show

wbMx.SSTab1.Tab = 1
End Sub


Private Sub txtCLF1_DblClick()
If txtKhmc.Text <> "" Then
    wbMx.Show
    wbMx.SSTab1.Tab = 1
    If Val(txtHtze.Text) > 0 Then
    wbMx.SSTab1.TabEnabled(3) = True
    Else
    wbMx.SSTab1.TabEnabled(3) = False
    End If
End If
wbMx.SSTab1.Enabled = True
End Sub


Private Sub txtCLF2_DblClick()
wbMx.Show

wbMx.SSTab1.Tab = 1
If Val(txtHtze.Text) > 0 Then
wbMx.SSTab1.TabEnabled(3) = True
Else
wbMx.SSTab1.TabEnabled(3) = False
End If
End Sub


Private Sub txtFkBz_DblClick()
If Val(txtHtze.Text) > 0 Then
wbMx.Show
wbMx.SSTab1.Tab = 3
End If
End Sub

Private Sub txtFbje1_Change()
''����ɱ��ܶ�
'txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
'Val(txtYf1.Text) + Val(txtQt1.Text)
End Sub

Private Sub txtGLG_Change()
'If txtGLG.Text <> "" And txtKhmc.Text <> "" Then
'cmdWb.Enabled = True
'Else
'cmdWb.Enabled = False
'End If
End Sub

Private Sub txtHtze_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = 13 Then
'����ɱ��ܶ�
txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
Val(txtYf1.Text) + Val(txtQt1.Text)

'��������
'If Val(txtHtze.Text) >= tmpHtze Then
    If optLa.Value = True Or optLb.Value = True Then
        txtJlr1.Text = Round(Val(txtHtze.Text) / 1.17 - Val(txtCbze1.Text), 2)
        txtLr1.Text = Round(Val(txtHtze.Text) / 1.17 - Val(txtCbze1.Text) - Val(txtYj1.Text), 2)
    ElseIf optLc.Value = True Then
        txtJlr1.Text = Round(Val(txtHtze.Text) / 1.06 - Val(txtCbze1.Text), 2)
        txtLr1.Text = Round(Val(txtHtze.Text) / 1.06 - Val(txtCbze1.Text) - Val(txtYj1.Text), 2)
    End If
wbMx.lblHtze.Caption = txtHtze.Text
'����Ӧ�ձ��еĽ��
frmFuK.adoHpt.Recordset.MoveFirst
Do While Not frmFuK.adoHpt.Recordset.EOF
frmFuK.adoHpt.Recordset.Fields("yingfJe").Value = frmFuK.adoHpt.Recordset.Fields("ED") _
* Val(wbMx.lblHtze.Caption)
frmFuK.adoHpt.Recordset.MoveNext
Loop
'Else
'txtHtze.Text = tmpHtze
'End If
End If
End Sub


Private Sub txtQt1_Change()
''����ɱ��ܶ�
'txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
'Val(txtYf1.Text) + Val(txtQt1.Text)
End Sub

Private Sub txtRgf1_DblClick()

If txtKhmc.Text <> "" Then

    wbMx.Show
    
    'wbMx.SSTab1.Caption = "�˹�����ϸ"
    wbMx.SSTab1.Tab = 0
    If Val(txtHtze.Text) > 0 Then
    wbMx.SSTab1.TabEnabled(3) = True
    Else
    wbMx.SSTab1.TabEnabled(3) = False
    End If
End If
wbMx.SSTab1.Enabled = True
End Sub

Private Sub txtRgf2_DblClick()
wbMx.Show

wbMx.SSTab1.Tab = 0
If Val(txtHtze.Text) > 0 Then
wbMx.SSTab1.TabEnabled(3) = True
Else
wbMx.SSTab1.TabEnabled(3) = False
End If
End Sub


Private Sub txtTc2_Change()
cmdSave.Enabled = True
End Sub

Private Sub txtYf1_Change()
''����ɱ��ܶ�
'txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
'Val(txtYf1.Text) + Val(txtQt1.Text)
End Sub

Private Sub txtYj1_Change()
''����ɱ��ܶ�
'txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
'Val(txtYf1.Text) + Val(txtQt1.Text)
'
''��������
'    If optLa.Value = True Or optLb.Value = True Then
'        txtJlr1.Text = Round(Val(txtHtze.Text) / 1.17 - Val(txtCbze1.Text), 2)
'        txtLr1.Text = Round(Val(txtHtze.Text) / 1.17 - Val(txtCbze1.Text) - Val(txtYj1.Text), 2)
'    ElseIf optLc.Value = True Then
'        txtJlr1.Text = Round(Val(txtHtze.Text) / 1.06 - Val(txtCbze1.Text), 2)
'        txtLr1.Text = Round(Val(txtHtze.Text) / 1.06 - Val(txtCbze1.Text) - Val(txtYj1.Text), 2)
'    End If
End Sub

Private Sub txtYj1_DblClick()

If mod1.DName = "����" Then
    If Val(txtYj2.Text) > 0 And frmYj.adoYj.Recordset.RecordCount = 0 Then
        MsgBox "�¾ɰ潻�浼����������,������������ϵ!"
        Exit Sub
    End If
    frmYj.cmdAdd.Visible = True
    frmYj.cmdDel.Visible = True
    frmYj.cmdSave.Visible = True
Else
    frmYj.cmdAdd.Visible = False
    frmYj.cmdAdd.Visible = False
    frmYj.cmdSave.Visible = False
End If
If mod1.DName = "������" Then
    frmYj.cmdAdd.Visible = True
    frmYj.cmdDel.Visible = True
    frmYj.cmdSave.Visible = True
End If
frmYj.Show
frmYj.lblHtbh.Caption = txtHtbh.Text
frmYj.lblKhmc.Caption = txtKhmc.Text
End Sub

Private Sub txtYj1_LostFocus()
'����ɱ��ܶ�
txtCbze1.Text = Val(txtClcb1.Text) + Val(txtRgf1.Text) + Val(txtCLF1.Text) + Val(txtFbje1.Text) + _
Val(txtYf1.Text) + Val(txtYj1.Text) + Val(txtQt1.Text)

'��������
If Val(txtHtze.Text) >= tmpHtze Then
    If optLa.Value = True Or optLb.Value = True Then
        txtLr1.Text = Round(Val(txtHtze.Text) / 1.17 - Val(txtCbze1.Text), 2)
        txtJlr1.Text = Round(Val(txtLr1.Text) + Val(txtYj1.Text), 2)
    ElseIf optLc.Value = True Then
        txtLr1.Text = Round(Val(txtHtze.Text) / 1.06 - Val(txtCbze1.Text), 2)
        txtJlr1.Text = Round(Val(txtLr1.Text) + Val(txtYj1.Text), 2)
    End If
End If
End Sub

Private Sub txtYj2_DblClick()
'frmYj.Show
'frmYj.lblHtbh.Caption = txtHtbh.Text
'frmYj.lblKhmc.Caption = txtKhmc.Text
End Sub


