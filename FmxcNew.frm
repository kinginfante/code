VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form FmxcNew 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ͬ����"
   ClientHeight    =   9090
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdBJ 
      BackColor       =   &H00C0FFFF&
      Caption         =   "  ����  �嵥"
      Height          =   765
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   8280
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgFk 
      Height          =   2745
      Left            =   240
      TabIndex        =   40
      Top             =   4320
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   4842
      _Version        =   393216
      BackColor       =   16777152
      Rows            =   30
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   12648384
      BackColorBkg    =   16777152
      FillStyle       =   1
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame frmFk 
      BackColor       =   &H00FFFFC0&
      Height          =   1215
      Left            =   270
      TabIndex        =   48
      Top             =   6960
      Width           =   4275
      Begin VB.CommandButton cmdGx 
         Caption         =   "����"
         Height          =   255
         Left            =   3030
         TabIndex        =   57
         Top             =   810
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CheckBox chkFP 
         BackColor       =   &H00C0FFC0&
         Caption         =   "��������,�����տ�"
         Height          =   255
         Left            =   1200
         TabIndex        =   128
         Top             =   870
         Width           =   1935
      End
      Begin VB.CheckBox chkKDFH 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�����"
         Height          =   285
         Left            =   60
         TabIndex        =   113
         Top             =   870
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "���"
         Height          =   255
         Left            =   3030
         TabIndex        =   56
         Top             =   150
         Width           =   825
      End
      Begin VB.CommandButton cmdDe 
         Caption         =   "ɾ��"
         Height          =   255
         Left            =   3030
         TabIndex        =   55
         Top             =   480
         Width           =   825
      End
      Begin VB.TextBox txtYrq 
         Height          =   300
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   150
         Width           =   1425
      End
      Begin VB.TextBox txtYje 
         Height          =   285
         Left            =   900
         TabIndex        =   49
         Top             =   480
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker dtpYf 
         Height          =   315
         Left            =   900
         TabIndex        =   51
         Top             =   150
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   16711680
         CalendarTrailingForeColor=   8454016
         Format          =   138084353
         CurrentDate     =   38797
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   735
         Index           =   1
         Left            =   0
         Top             =   90
         Width           =   2775
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Ӧ������"
         Height          =   285
         Left            =   60
         TabIndex        =   54
         Top             =   180
         Width           =   735
      End
      Begin VB.Label lblFid 
         BackStyle       =   0  'Transparent
         Caption         =   "lblFid"
         Height          =   165
         Left            =   2760
         TabIndex        =   53
         Top             =   840
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "�տ���"
         Height          =   225
         Left            =   60
         TabIndex        =   52
         Top             =   570
         Width           =   795
      End
   End
   Begin VB.TextBox txtCompanyId 
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   90
      Text            =   "Text1"
      Top             =   8280
      Width           =   3105
   End
   Begin VB.TextBox txtHtbh 
      Height          =   270
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1395
      Width           =   3345
   End
   Begin VB.Frame frmFP 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   3720
      TabIndex        =   129
      Top             =   4800
      Width           =   735
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "��������  �����տ�"
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Left            =   240
         TabIndex        =   130
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame frmYG 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   765
      Left            =   9180
      TabIndex        =   120
      Top             =   3720
      Width           =   1965
   End
   Begin VB.TextBox txtQB1 
      Height          =   285
      Left            =   10290
      TabIndex        =   127
      Text            =   "Text2"
      Top             =   4080
      Width           =   795
   End
   Begin VB.CommandButton cmdYongYou 
      Caption         =   "��������"
      Height          =   735
      Left            =   11400
      TabIndex        =   126
      Top             =   7320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "������"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   4560
      TabIndex        =   77
      Top             =   7320
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   1200
         TabIndex        =   132
         Top             =   120
         Width           =   3495
         Begin VB.OptionButton optG2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "��ԭ��"
            Height          =   255
            Left            =   1320
            TabIndex        =   134
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton optG1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "����"
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   120
            Width           =   975
         End
      End
      Begin NiceFormControl.NiceButton NiceButton1 
         Height          =   945
         Left            =   5220
         TabIndex        =   104
         Top             =   330
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1667
         BTYPE           =   3
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16761024
         BCOLO           =   16761024
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FmxcNew.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Style           =   21
         Caption         =   "����"
      End
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1065
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   81
         Top             =   540
         Width           =   4965
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ͬ��"
         Height          =   225
         Left            =   5220
         TabIndex        =   80
         Top             =   510
         Width           =   705
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�ܾ�"
         Height          =   195
         Left            =   5220
         TabIndex        =   79
         Top             =   870
         Width           =   675
      End
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "����"
         Height          =   285
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame frmNewLx 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   4920
      TabIndex        =   116
      Top             =   1320
      Width           =   10035
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgNLN 
         Height          =   255
         Left            =   6660
         TabIndex        =   119
         Top             =   3330
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Cols            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.Frame frmTJ 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   0
         TabIndex        =   118
         Top             =   3330
         Width           =   7125
         Begin VB.OptionButton optAb 
            BackColor       =   &H00FFFFC0&
            Caption         =   "׷�ӵ�"
            Enabled         =   0   'False
            Height          =   180
            Left            =   3240
            TabIndex        =   124
            Top             =   60
            Width           =   945
         End
         Begin VB.OptionButton OptAc 
            BackColor       =   &H00FFFFC0&
            Caption         =   "���ջ���"
            Height          =   255
            Left            =   4200
            TabIndex        =   123
            Top             =   30
            Width           =   1065
         End
         Begin VB.OptionButton optAA 
            BackColor       =   &H00FFFFC0&
            Caption         =   "ѯ�۵�"
            Enabled         =   0   'False
            Height          =   225
            Left            =   2280
            TabIndex        =   122
            Top             =   60
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.TextBox txtFX 
            Height          =   285
            Left            =   5340
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   30
            Width           =   1395
         End
         Begin NiceFormControl.NiceButton cmdTj 
            Height          =   345
            Left            =   120
            TabIndex        =   125
            Top             =   0
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   609
            BTYPE           =   3
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
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
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FmxcNew.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
            Caption         =   "���ҵ����Ŀ"
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgNewLx 
         Height          =   3255
         Left            =   30
         TabIndex        =   117
         Top             =   60
         Width           =   9345
         _ExtentX        =   16484
         _ExtentY        =   5741
         _Version        =   393216
         BackColor       =   16777152
         Rows            =   14
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   16777152
         BackColorBkg    =   16777152
         AllowUserResizing=   1
         BandDisplay     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
   End
   Begin VB.ComboBox comQBF 
      Height          =   300
      ItemData        =   "FmxcNew.frx":0038
      Left            =   9330
      List            =   "FmxcNew.frx":0042
      TabIndex        =   115
      Top             =   4080
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgRen 
      Height          =   1275
      Left            =   10170
      TabIndex        =   109
      Top             =   4650
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2249
      _Version        =   393216
      BackColor       =   16777152
      Rows            =   10
      FixedCols       =   0
      BackColorFixed  =   12648384
      BackColorBkg    =   16777152
      SelectionMode   =   1
      BorderStyle     =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox Text1 
      Height          =   1005
      Left            =   12300
      Locked          =   -1  'True
      TabIndex        =   110
      Text            =   "Text1"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtZBZ 
      BackColor       =   &H00C0FFC0&
      Height          =   795
      Left            =   10920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   108
      Text            =   "FmxcNew.frx":0054
      ToolTipText     =   "�˴����ۺϹ�����д"
      Top             =   2850
      Width           =   4185
   End
   Begin VB.TextBox txtQb 
      Height          =   300
      Left            =   10290
      TabIndex        =   106
      Text            =   "Text1"
      Top             =   4080
      Width           =   825
   End
   Begin VB.Frame frmYm 
      BackColor       =   &H00FFFFC0&
      Caption         =   "��Ŀ������ϸ:"
      ForeColor       =   &H000000FF&
      Height          =   2265
      Left            =   7830
      TabIndex        =   95
      Top             =   5760
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdYdel 
         Caption         =   "ɾ��"
         Height          =   285
         Left            =   3990
         TabIndex        =   100
         Top             =   1290
         Width           =   585
      End
      Begin VB.CommandButton cmdYadd 
         Caption         =   "���"
         Height          =   315
         Left            =   3990
         TabIndex        =   99
         Top             =   930
         Width           =   585
      End
      Begin VB.TextBox txtYingFu 
         Height          =   270
         Left            =   2880
         TabIndex        =   98
         Top             =   1710
         Width           =   1035
      End
      Begin VB.TextBox txtFED 
         Height          =   285
         Left            =   960
         TabIndex        =   97
         Top             =   1710
         Width           =   645
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "�ر�"
         Height          =   285
         Left            =   3990
         TabIndex        =   96
         Top             =   1680
         Width           =   615
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgYJ 
         Height          =   1275
         Left            =   150
         TabIndex        =   101
         Top             =   300
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   2249
         _Version        =   393216
         BackColor       =   16777152
         Rows            =   10
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   12648384
         BackColorBkg    =   16777152
         SelectionMode   =   1
         BorderStyle     =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   315
         Left            =   1680
         TabIndex        =   107
         Top             =   1740
         Width           =   105
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "֧�����"
         Height          =   225
         Left            =   2010
         TabIndex        =   103
         Top             =   1740
         Width           =   915
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "�տ���"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   1740
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�޸�"
      Height          =   765
      Left            =   12660
      Picture         =   "FmxcNew.frx":005A
      Style           =   1  'Graphical
      TabIndex        =   92
      ToolTipText     =   "�޸�"
      Top             =   8280
      Width           =   675
   End
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "���"
      Height          =   765
      Left            =   11970
      Picture         =   "FmxcNew.frx":0364
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   8280
      Width           =   675
   End
   Begin VB.ComboBox companyId 
      Height          =   300
      ItemData        =   "FmxcNew.frx":07A6
      Left            =   1380
      List            =   "FmxcNew.frx":07B6
      TabIndex        =   89
      Text            =   "�Ϻ���������յ��������޹�˾"
      Top             =   8280
      Width           =   3375
   End
   Begin VB.Frame frmYJ 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   12390
      TabIndex        =   85
      Top             =   5700
      Width           =   2595
      Begin VB.TextBox txtYjBz 
         Height          =   915
         Left            =   1050
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   94
         Text            =   "FmxcNew.frx":0826
         Top             =   630
         Width           =   1305
      End
      Begin VB.TextBox txtYJ 
         Height          =   270
         Left            =   1050
         TabIndex        =   86
         Text            =   "Text1"
         Top             =   90
         Width           =   1305
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "��ע"
         Height          =   195
         Left            =   480
         TabIndex        =   93
         Top             =   630
         Width           =   525
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   195
         Left            =   480
         TabIndex        =   87
         Top             =   150
         Width           =   855
      End
   End
   Begin NiceFormControl.NiceCheck optYj 
      Height          =   195
      Left            =   12900
      TabIndex        =   84
      Top             =   6330
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "��Ŀ����"
      BackColor       =   16777152
   End
   Begin VB.TextBox txtBz 
      BackColor       =   &H00FFFFC0&
      Height          =   2835
      Left            =   10890
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   75
      Text            =   "FmxcNew.frx":082C
      Top             =   30
      Width           =   4215
   End
   Begin NiceFormControl.NiceOption optLx 
      Height          =   240
      Left            =   11280
      TabIndex        =   74
      Top             =   3840
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Caption         =   "ҵ������˵��"
      BackColor       =   16777152
   End
   Begin NiceFormControl.NiceOption optXm 
      Height          =   240
      Left            =   11250
      TabIndex        =   73
      Top             =   4140
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "��Ŀ��ע"
      BackColor       =   16777152
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgFKN 
      Bindings        =   "FmxcNew.frx":0837
      Height          =   855
      Left            =   5100
      TabIndex        =   59
      Top             =   6360
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   1508
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   3150
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   2610
   End
   Begin NiceFormControl.NiceButton cmdKQy 
      Height          =   345
      Left            =   2760
      TabIndex        =   47
      Top             =   2640
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   609
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FmxcNew.frx":084D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Caption         =   "��������"
   End
   Begin VB.TextBox txtFPLx 
      Height          =   270
      Left            =   3510
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "Text1"
      Top             =   3570
      Width           =   1245
   End
   Begin VB.TextBox txtEd 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00008000&
      Height          =   270
      Left            =   3510
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   4020
      Width           =   885
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����"
      Height          =   765
      Left            =   13350
      Picture         =   "FmxcNew.frx":0869
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8280
      Width           =   675
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����"
      Height          =   765
      Left            =   14700
      Picture         =   "FmxcNew.frx":0ED3
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8280
      Width           =   585
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   765
      Left            =   14040
      Picture         =   "FmxcNew.frx":0FD5
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8280
      Width           =   645
   End
   Begin VB.TextBox txtQy 
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   2640
      Width           =   1155
   End
   Begin VB.TextBox txtBm 
      Height          =   270
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   3060
      Width           =   1155
   End
   Begin VB.TextBox txtHtxz 
      Height          =   270
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   1830
      Width           =   1155
   End
   Begin VB.ComboBox comFPLX 
      Height          =   300
      ItemData        =   "FmxcNew.frx":115F
      Left            =   3510
      List            =   "FmxcNew.frx":116F
      TabIndex        =   35
      Text            =   "Combo1"
      Top             =   3570
      Width           =   1545
   End
   Begin VB.TextBox txtF 
      Height          =   300
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   4050
      Width           =   1425
   End
   Begin VB.TextBox txtL 
      Height          =   300
      Left            =   7470
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   4050
      Width           =   1305
   End
   Begin VB.TextBox txtKhmc 
      Height          =   270
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   975
      Width           =   3345
   End
   Begin VB.TextBox txtYwy 
      Height          =   270
      Left            =   3510
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   8760
      Width           =   945
   End
   Begin VB.TextBox txtXMMC 
      Height          =   270
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   540
      Width           =   3345
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
      Height          =   270
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "���ڸ�����ϸ��ȷ����ͬ�ܽ��"
      Top             =   3570
      Width           =   1125
   End
   Begin VB.TextBox txtXYwy 
      Height          =   270
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   8745
      Width           =   1035
   End
   Begin VB.TextBox txtZe 
      ForeColor       =   &H00008000&
      Height          =   270
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4020
      Width           =   1125
   End
   Begin VB.TextBox txtHtrq 
      Height          =   270
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   3345
   End
   Begin VB.TextBox txtYjpw 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   11670
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   7740
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.TextBox txtZbh 
      Height          =   270
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2280
      Width           =   2265
   End
   Begin VB.CommandButton cmdHt 
      BackColor       =   &H008080FF&
      Caption         =   "BH"
      Height          =   225
      Left            =   1050
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   315
   End
   Begin MSComCtl2.DTPicker dt4 
      Height          =   315
      Left            =   7470
      TabIndex        =   27
      Top             =   4050
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy��M��d��"
      Format          =   138412035
      CurrentDate     =   38098
   End
   Begin MSComCtl2.DTPicker dt3 
      Height          =   315
      Left            =   5280
      TabIndex        =   28
      Top             =   4050
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy��M��d��"
      Format          =   138412035
      CurrentDate     =   38098
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLx 
      Height          =   3645
      Left            =   5100
      TabIndex        =   41
      Top             =   30
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   6429
      _Version        =   393216
      BackColor       =   16777152
      Rows            =   14
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   16777152
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin MSComDlg.CommonDialog cmdDia 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoFile 
      Height          =   375
      Left            =   0
      Top             =   480
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
   Begin NiceFormControl.NiceButton cmdDz 
      Height          =   345
      Left            =   2670
      TabIndex        =   76
      Top             =   1800
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   609
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FmxcNew.frx":1199
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Caption         =   "���Ӻ�ͬ"
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   4515
      Left            =   4830
      TabIndex        =   82
      Top             =   4530
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   7964
      _Version        =   393216
      BackColor       =   14414066
      ForeColor       =   8404992
      Rows            =   15
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   16761024
      ForeColorFixed  =   0
      BackColorBkg    =   14414066
      GridColorFixed  =   8404992
      GridColorUnpopulated=   8404992
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin NiceFormControl.NiceButton cmdDz1 
      Height          =   345
      Left            =   3750
      TabIndex        =   112
      Top             =   1800
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   609
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FmxcNew.frx":11B5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Caption         =   "��������"
   End
   Begin NiceFormControl.NiceButton cmdZX 
      Height          =   345
      Left            =   3720
      TabIndex        =   135
      Top             =   2280
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   609
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FmxcNew.frx":11D1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Caption         =   "ִ��״��"
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   405
      Index           =   5
      Left            =   120
      Top             =   8700
      Width           =   2565
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "ȫ����"
      Height          =   255
      Left            =   9360
      TabIndex        =   114
      Top             =   3810
      Width           =   555
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFC0&
      Caption         =   "��ϵ��"
      Height          =   225
      Left            =   12360
      TabIndex        =   111
      Top             =   4410
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Ԥ���ɱ�"
      Height          =   165
      Left            =   10260
      TabIndex        =   105
      Top             =   3810
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   525
      Index           =   4
      Left            =   2700
      Top             =   3420
      Width           =   2385
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   705
      Index           =   3
      Left            =   5130
      Top             =   3750
      Width           =   4005
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   405
      Index           =   2
      Left            =   120
      Top             =   8250
      Width           =   4965
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   705
      Index           =   0
      Left            =   9270
      Top             =   3750
      Width           =   1845
   End
   Begin VB.Label lblCom 
      BackStyle       =   0  'Transparent
      Caption         =   "ǩԼ��˾"
      Height          =   225
      Left            =   390
      TabIndex        =   88
      Top             =   8340
      Width           =   825
   End
   Begin VB.Label lblTX 
      BackStyle       =   0  'Transparent
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
      Height          =   1125
      Left            =   11790
      TabIndex        =   83
      Top             =   7140
      Width           =   3315
   End
   Begin VB.Label lblMy 
      BackStyle       =   0  'Transparent
      Caption         =   "Label19"
      Height          =   195
      Left            =   13920
      TabIndex        =   72
      Top             =   5616
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "ó��"
      Height          =   195
      Left            =   12900
      TabIndex        =   71
      Top             =   5620
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblLR 
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      Height          =   195
      Left            =   13920
      TabIndex        =   70
      Top             =   5970
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblZJ 
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      Height          =   195
      Left            =   13920
      TabIndex        =   69
      Top             =   5262
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblYs 
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      Height          =   195
      Left            =   13920
      TabIndex        =   68
      Top             =   4908
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblRGF 
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      Height          =   195
      Left            =   13920
      TabIndex        =   67
      Top             =   4554
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   195
      Left            =   12900
      TabIndex        =   66
      Top             =   5975
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "�н�"
      Height          =   195
      Left            =   12900
      TabIndex        =   65
      Top             =   5265
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "ѹ����"
      Height          =   195
      Left            =   12900
      TabIndex        =   64
      Top             =   4910
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "�˹�"
      Height          =   195
      Left            =   12900
      TabIndex        =   63
      Top             =   4555
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblCBZE 
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      Height          =   195
      Left            =   13920
      TabIndex        =   62
      Top             =   4200
      Width           =   945
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "�ɱ��ܶ�"
      Height          =   195
      Left            =   12900
      TabIndex        =   61
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "M F ϵ��"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   12900
      TabIndex        =   60
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblHid 
      Caption         =   "lblHid"
      Height          =   285
      Left            =   11070
      TabIndex        =   58
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ʊ����"
      Height          =   255
      Left            =   2700
      TabIndex        =   36
      Top             =   3630
      Width           =   975
   End
   Begin VB.Label lblHTF 
      BackStyle       =   0  'Transparent
      Caption         =   "״̬"
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   3570
      TabIndex        =   34
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "ִ��״̬"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2730
      TabIndex        =   33
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "��  ��"
      Height          =   255
      Left            =   300
      TabIndex        =   32
      Top             =   3120
      Width           =   705
   End
   Begin VB.Label Label52 
      BackStyle       =   0  'Transparent
      Caption         =   "ά��������"
      Height          =   225
      Left            =   7590
      TabIndex        =   31
      Top             =   3780
      Width           =   1275
   End
   Begin VB.Label Label51 
      BackStyle       =   0  'Transparent
      Caption         =   "ά����ʼ��"
      Height          =   225
      Left            =   5490
      TabIndex        =   30
      Top             =   3810
      Width           =   1605
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "---��"
      Height          =   225
      Left            =   7020
      TabIndex        =   29
      Top             =   4110
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ִ�б��"
      Height          =   255
      Left            =   300
      TabIndex        =   24
      Top             =   2330
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "��ͬ����"
      Height          =   255
      Left            =   300
      TabIndex        =   23
      Top             =   1897
      Width           =   975
   End
   Begin VB.Label lblhtbh 
      BackStyle       =   0  'Transparent
      Caption         =   "��ͬ���"
      Height          =   255
      Left            =   300
      TabIndex        =   22
      Top             =   1464
      Width           =   975
   End
   Begin VB.Label lblKhmc 
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ�����"
      Height          =   255
      Left            =   300
      TabIndex        =   20
      Top             =   1031
      Width           =   975
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "�ͷ�����"
      Height          =   255
      Left            =   2730
      TabIndex        =   19
      Top             =   8820
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ����"
      Height          =   255
      Left            =   300
      TabIndex        =   18
      Top             =   598
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "��ͬ���"
      Height          =   255
      Left            =   300
      TabIndex        =   17
      Top             =   3630
      Width           =   975
   End
   Begin VB.Label lblHtrq 
      BackStyle       =   0  'Transparent
      Caption         =   "��ͬ����"
      Height          =   255
      Left            =   300
      TabIndex        =   16
      Top             =   165
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "���ؾ���"
      Height          =   255
      Index           =   0
      Left            =   390
      TabIndex        =   15
      Top             =   8790
      Width           =   945
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Caption         =   "��  ��"
      Height          =   255
      Left            =   300
      TabIndex        =   14
      Top             =   2715
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "�տ���"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2700
      TabIndex        =   13
      Top             =   4050
      Width           =   915
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "ʵ���տ�"
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   300
      TabIndex        =   12
      Top             =   4080
      Width           =   795
   End
   Begin VB.Label Label50 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   4530
      TabIndex        =   11
      Top             =   4050
      Width           =   195
   End
   Begin VB.Label lblMF 
      BackStyle       =   0  'Transparent
      Caption         =   "MF"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   13950
      TabIndex        =   10
      Top             =   3840
      Width           =   2115
   End
End
Attribute VB_Name = "FmxcNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lc As Integer
Dim Fwid As Long
Dim LCRen As String
Dim LCUid As String
Dim timZm As Integer

Dim W1 As Single 'ά����׼��
Dim W2 As Single '����
Dim W3 As Single '�����˹�
Dim W4 As Single 'ѹ����ά�ޱ���
Dim W5 As Single 'ѹ����ó��
Dim W6 As Single '�н�ҵ��
Dim W7 As Single '����
Dim W8 As Single '����
Dim W9 As Single '�ڴ︻
Dim W10 As Single '��ͼ
Dim W11 As Single '�����
Dim W12 As Single '�ְ�
Dim W13 As Single '�Ǵ����Ʒ

Dim D1 As Double '�ٴ���
Dim D2 As Double
Dim D3 As Double
Dim D4 As Double
Dim D5 As Double
Dim D6 As Double
Dim D7 As Double
Dim D8 As Double
Dim D9 As Double
Dim D10 As Double
Dim D11 As Double
Dim D12 As Double
Dim D13 As Double
Public XJZL As String '����fmxcxj��lblZl

Dim LLXX As Boolean '(�½��˹�ѯ�ۣ��������ѯ�ۣ�
Public NewId As Integer '�½�ѯ�۵����к�

Dim YGCB As Double 'Ԥ���ɱ����ڱ���ʱ����Ԥ��ѯ�۵��ϼ��

Dim QBZE As Double 'Ԥ���ɱ��ܶ�,�ٴ���
Public Bid As Long
Public HTLX As String

Public Sub DJ() '�����ٴ���
On Error Resume Next
Dim CB As Single
Dim ZE As Single

'�����ٴ���
CB = Val(lblCBZE.Caption)
ZE = Val(txtHtze.Text)
If W1 > 0 Then
    If CB - W1 = 0 Then
        D1 = ZE
    Else
        D1 = Round((ZE * W1) / CB, 2)
        D1 = Round(D1, 2)
    End If
End If
If W2 > 0 Then
    If CB - (W2 + W1) = 0 Then
        D2 = Round((ZE - D1), 2)
    Else
        D2 = Round(ZE * W2 / CB, 2)
        D2 = Round(D2, 2)
    End If
End If
If W3 > 0 Then
    If CB - (W3 + W1 + W2) = 0 Then
        D3 = Round((ZE - D1 - D2), 2)
    Else
        D3 = Round(ZE * W3 / CB, 2)
        D3 = Round(D3, 2)
    End If
End If
If W4 > 0 Then
    If CB - (W4 + W3 + W1 + W2) = 0 Then
        D4 = Round((ZE - D1 - D2 - D3), 2)
    Else
        D4 = Round(ZE * W4 / CB, 2)
        D4 = Round(D4, 2)
    End If
End If
If W5 > 0 Then
    If CB - (W5 + W4 + W3 + W1 + W2) = 0 Then
        D5 = Round((ZE - D1 - D2 - D3 - D4), 2)
    Else
        D5 = Round(ZE * W5 / CB, 2)
        D5 = Round(D5, 2)
    End If
End If
If W6 > 0 Then
    If CB - (W6 + W5 + W4 + W3 + W1 + W2) = 0 Then
        D6 = Round((ZE - D1 - D2 - D3 - D4 - D5), 2)
    Else
        D6 = Round(ZE * W6 / CB, 2)
        D6 = Round(D6, 2)
    End If
End If
If W7 > 0 Then
    If CB - (W7 + W6 + W5 + W4 + W3 + W1 + W2) = 0 Then
        D7 = Round((ZE - D1 - D2 - D3 - D4 - D5 - D6), 2)
    Else
        D7 = Round(ZE * W7 / CB, 2)
        D7 = Round(D7, 2)
    End If
End If
If W8 > 0 Then
    If CB - (W8 + W7 + W6 + W5 + W4 + W3 + W1 + W2) = 0 Then
        D8 = Round((ZE - D1 - D2 - D3 - D4 - D5 - D6 - D7), 2)
    Else
        D8 = Round(ZE * W8 / CB, 2)
        D8 = Round(D8, 2)
    End If
End If
If W9 > 0 Then
    If CB - (W9 + W8 + W7 + W6 + W5 + W4 + W3 + W1 + W2) = 0 Then
        D9 = Round((ZE - D1 - D2 - D3 - D4 - D5 - D6 - D7 - D8), 2)
    Else
        D9 = Round(ZE * W9 / CB, 2)
        D9 = Round(D9, 2)
    End If
End If
If W10 > 0 Then
    If CB - (W10 + W9 + W8 + W7 + W6 + W5 + W4 + W3 + W1 + W2) = 0 Then
        D10 = Round((ZE - D1 - D2 - D3 - D4 - D5 - D6 - D7 - D8 - D9), 2)
    Else
        D10 = Round(ZE * W10 / CB, 2)
        D10 = Round(D10, 2)
    End If
End If
If W11 > 0 Then
    If CB - (W11 + W10 + W9 + W8 + W7 + W6 + W5 + W4 + W3 + W1 + W2) = 0 Then
        D11 = Round((ZE - D1 - D2 - D3 - D4 - D5 - D6 - D7 - D8 - D9 - D10), 2)
    Else
        D11 = Round(ZE * W11 / CB, 2)
        D11 = Round(D11, 2)
    End If
End If
If W12 > 0 Then
    If CB - (W12 + W11 + W10 + W9 + W8 + W7 + W6 + W5 + W4 + W3 + W1 + W2) = 0 Then
        D12 = Round((ZE - D1 - D2 - D3 - D4 - D5 - D6 - D7 - D8 - D9 - D10 - D11), 2)
    Else
        D12 = Round(ZE * W12 / CB, 2)
        D12 = Round(D12, 2)
    End If
End If
If W13 > 0 Then
    If CB - (W13 + W12 + W11 + W10 + W9 + W8 + W7 + W6 + W5 + W4 + W3 + W1 + W2) = 0 Then
        D13 = Round((ZE - D1 - D2 - D3 - D4 - D5 - D6 - D7 - D8 - D9 - D10 - D11 - D12), 2)
    Else
        D13 = Round(ZE * W13 / CB, 2)
        D13 = Round(D13, 2)
    End If
End If
dtgLx.Col = 3
dtgLx.Row = 1: If D1 > 0 Then dtgLx.Text = D1
dtgLx.Row = 2: If D2 > 0 Then dtgLx.Text = D2
dtgLx.Row = 3: If D3 > 0 Then dtgLx.Text = D3
dtgLx.Row = 4: If D4 > 0 Then dtgLx.Text = D4
dtgLx.Row = 5: If D5 > 0 Then dtgLx.Text = D5
dtgLx.Row = 6: If D6 > 0 Then dtgLx.Text = D6
dtgLx.Row = 7: If D7 > 0 Then dtgLx.Text = D7
dtgLx.Row = 8: If D8 > 0 Then dtgLx.Text = D8
dtgLx.Row = 9: If D9 > 0 Then dtgLx.Text = D9
dtgLx.Row = 10: If D10 > 0 Then dtgLx.Text = D10
dtgLx.Row = 11: If D11 > 0 Then dtgLx.Text = D11
dtgLx.Row = 12: If D12 > 0 Then dtgLx.Text = D12
dtgLx.Row = 13: If D13 > 0 Then dtgLx.Text = D13
End Sub

Private Sub cmdAdd_Click()
timZm = 1 '����༭
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
    mod1.cmd.Parameters("@NBLX") = "����༭"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "���"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtYje.Text)
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = chkKDFH.Value '�����
    mod1.cmd.Parameters("@md1") = txtYRQ.Text
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

'''''        frmFX.Visible = False
        
    End If

    
Set mod1.cmd = Nothing
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
If htBrow.Visible = True Then
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

Private Sub cmdBJ_Click()
FmxcBJ.Show
Call FmxcBJ.dtgbrFF
Call FmxcBJ.Bound(Val(lblHid.Caption))

End Sub

Private Sub cmdClose_Click()
frmYm.Visible = False
End Sub

Private Sub cmdDe_Click()
Dim ii As Integer
ii = MsgBox("�Ƿ�ɾ���˱ʸ����¼?", vbYesNo + vbQuestion, "����")
If ii = vbNo Then Exit Sub

timZm = 1 '����༭
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
    mod1.cmd.Parameters("@NBLX") = "����༭"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "ɾ��"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtYje.Text)
    mod1.cmd.Parameters("@mm20") = Val(lblFid.Caption)
    mod1.cmd.Parameters("@mb1") = Null
    
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


        
    End If

    
Set mod1.cmd = Nothing
End Sub

Private Sub cmdDel_Click()
Dim ii As Integer
Dim YY As String
If txtHtbh.Text <> "HMNEW" Then
    If mod1.DName <> "������" And mod1.DName <> "�Ǽ���" And mod1.DName <> "����" And mod1.DName <> "�Ǽ���" And mod1.DName <> txtYwy.Text Then
         Exit Sub
    End If
    If Lc > 1 And mod1.DName = txtYwy.Text Then
        Exit Sub
    End If
End If
ii = MsgBox("�Ƿ����ϴ˺�ͬ���󵥣�", vbYesNo + vbQuestion, "Hello")
If ii = vbNo Then
    Exit Sub
End If
YY = InputBox("����������ԭ��!")

timZm = 12 'ɾ����ͬ
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "ɾ����ͬ"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mlt1") = YY
    mod1.cmd.Parameters("@mm1") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
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
Dim oo As Integer
Dim ii As Integer
Dim Ra
Dim tt As String
Dim TC As Integer '����
Dim Rf As Boolean
Dim YJF As Boolean '�Ƿ��������
Dim BZE As Single
On Error Resume Next
If mod1.GxName = "���۹���" And mod1.GXF = True And Me.HTLX = "ѯ��ָ��" Then
'����ͬ�ܶ��뱨�۷����ܶ��Ƿ�һ��
If OptT1.Value = True Then
    BZE = 0: dtgNLN.Col = 2
    For oo = 1 To 100
        dtgNLN.Row = oo
        dtgNLN.Col = 2
        BZE = BZE + Val(dtgNLN.Text)
        dtgNLN.Col = 0
        If dtgNLN.Text = "" Then Exit For
    Next
    If BZE <> Val(txtHtze.Text) Then
        ii = MsgBox("���۷����ܶ�Ϊ:" & BZE & ",���ͬ�ܶ�:" & txtHtze.Text & "��һ�£���ȷ�ϣ�", vbInformation + vbOKOnly, "��ע��")
        Exit Sub
    End If
End If
End If
If OptT1.Value = False And optT2.Value = False Then
    Exit Sub
End If
If OptT1.Value = True And Me.JCYG = True Then
    MsgBox ("��ͬ��ֻ����һ��Ԥ���ɱ�ѯ�۵���")
    Exit Sub
End If
If txtQy.Text <> "�Ϻ�" And (mod1.DName = "�߶���" Or mod1.DName = "�Ǽ���") Then
    If optG1.Value = False And optG2.Value = False Then
        MsgBox "��ȷ�ϸ��»����ջ�ԭ��!"
        Exit Sub
    End If
End If
tt = "select lc,lcuid,lcren,htf from htping where hid=" & Val(lblHid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing

If IsNull(Ra(0, 0)) = True Then
    MsgBox "������ϣ������ԣ����˳�������Ϣ���ԣ�"
    Exit Sub
End If
Lc = Ra(0, 0)
LCUid = Ra(1, 0)
LCRen = Ra(2, 0)
lblHTF.ToolTipText = Ra(3, 0)
Select Case lblHTF.ToolTipText
Case 0
    lblHTF.Caption = "�༭"
Case 6
    lblHTF.Caption = "����"
Case 9
    lblHTF.Caption = "����"
Case 1
    lblHTF.Caption = "��ִ��"
Case 2
    lblHTF.Caption = "���"
Case 3
    lblHTF.Caption = "ִ����"
Case 100
    lblHTF.Caption = "���"
End Select

If (mod1.DName = "�߶���" Or mod1.DName = "�Ǽ���" Or mod1.DName = txtYwy.Text) Then
    LCRen = mod1.DName: LCUid = mod1.DHid

End If

If Lc = 100 And mod1.DName <> "����" And mod1.DName <> "������" And txtQy.Text = "�Ϻ�" Then
    Exit Sub
End If
If LCUid <> mod1.DHid And OptT1.Value = True Then
    MsgBox "�˴�Ӧ��" & LCRen & "����! ������Ҫ�ٵ�"
    Exit Sub
End If


If Lc = 1 Then
    If txtHtbh.Text = "HMNEW" Then
        MsgBox ("�������ɺ�ͬ���!")
        Exit Sub
    End If
    
    If Val(cmdDZ.ToolTipText) = 0 And Lc = 1 Then
        MsgBox "�뵼����Ӱ��ͬ(����!"
        Call cmdDZ_Click
        frmQm.Visible = False
        Exit Sub
    End If
    
    '��������˹������⼼������
    Rf = False
    dtgLx.Col = 2
    For oo = 1 To 13
        dtgLx.Row = oo
        If Val(dtgLx.Text) > 0 And (oo = 1 Or oo = 2 Or oo = 3 Or oo = 4 Or oo = 12) Then
            Rf = True
            Exit For
        End If
    Next
    
    If Val(cmdDz1.ToolTipText) = 0 And Lc = 1 And Rf = True Then
        MsgBox "�뵼�뼼������!"
        Call cmdDz1_Click
        frmQm.Visible = False
        Exit Sub
    End If
    
    If txtFPLx.Text = "" Then
        Me.comFPLX.Visible = True
        MsgBox ("��ѡ��Ʊ��ʽ!")
        cmdSave.Enabled = True
        frmQm.Visible = False
        Exit Sub
    End If
    
    If W1 > 0 And (txtF.Text = "" Or txtL.Text = "") Then
        MsgBox ("�����ά������ʼ�ںͽ�����!")
        dt3.Visible = True: dt4.Visible = True
        cmdSave.Enabled = True
        frmQm.Visible = False
        Exit Sub
    End If
    
'''''    If optYj.Value = Mixed Then
'''''        MsgBox ("��ȷ���Ƿ������Ŀ����!")
'''''        Exit Sub
'''''    End If
    If txtYjBz.Text = "" And mod1.Qy = "�Ϻ�" Then
        MsgBox "������˿ͻ�����ϵ��!" & Chr(13) & Chr(10) & "(�˹������������ͻ����ϵ�¼��!������ȷ¼��ͻ���ϵ����Ϣ,���˫����ϵ�˵�����,����ɲ���)"
        Call txtKhmc_DblClick
        Exit Sub
    End If
End If

On Error Resume Next
If optT2.Value = True Then
    If txtYwy.Text <> mod1.DName And txtXYwy.Text <> mod1.DName And Trim(LCUid) <> mod1.DHid And mod1.DName <> "����" And mod1.DName <> "������" Then
        MsgBox "���ز�����ֻ���ɱ��˽��У�"
        Exit Sub
    End If
End If
If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("����һ��Ҫ���߾ܾ��ҵ�����!  :) ")
    Exit Sub
End If
If optT2.Value = True Then
    If (lblHTF.Caption = "ִ����" Or lblHTF.Caption = "���") And mod1.DName <> "���ĳ�" And mod1.DName <> "�Ǽ���" And mod1.DName <> "����" And mod1.Mname <> "������" Then
    MsgBox ("�Ѿ�ִ����ϵĺ�ͬ�����ܹ�����!  :) ")
    Exit Sub
    End If
End If
frmFX.Visible = False

If Val(txtYJ.Text) > 0 Then
    YJF = True
End If
If OptT1.Value = True And Lc > 1 And (mod1.DName = "֣��" Or mod1.DName = "�Ǽ���" Or mod1.DName = "�����" Or mod1.DName = "�ܴ���") And Val(txtYJ.Text) = 0 Then
    ii = MsgBox("�Ƿ��������?", vbQuestion + vbYesNo + vbDefaultButton2, "��ȷ��")
    If ii = vbYes Then
        YJF = True
        txtQM.Text = txtQM.Text & " Сֽ��"
    End If
End If
If OptT1.Value = True And Lc = 1 And (mod1.DName = "�ܴ���") And Val(txtYJ.Text) = 0 Then
    ii = MsgBox("�Ƿ��������?", vbQuestion + vbYesNo + vbDefaultButton2, "��ȷ��")
    If ii = vbYes Then
        YJF = True
        txtQM.Text = txtQM.Text & " Сֽ��"
    End If
End If
timZm = 10 'ǩ��
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
    mod1.cmd.Parameters("@NBLX") = "ǩ��"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtXYwy.Text
    mod1.cmd.Parameters("@mt2") = txtXYwy.ToolTipText
    mod1.cmd.Parameters("@mt3") = txtXmmc.Text
    mod1.cmd.Parameters("@mt4") = txtHtbh.Text
    
    mod1.cmd.Parameters("@mt5") = lblHTF.Caption '״̬
    mod1.cmd.Parameters("@mt6") = mod1.GJId
   
    mod1.cmd.Parameters("@mt15") = txtHtxz.Text
    mod1.cmd.Parameters("@mlt1") = txtQM.Text '������
    If mod1.Qy <> "�Ϻ�" And Me.HTLX = "ѯ��ָ��" Then
        Lc = 10
    End If
    mod1.cmd.Parameters("@mm1") = Lc

    mod1.cmd.Parameters("@mm2") = Fwid
    mod1.cmd.Parameters("@mm10") = Val(txtHtze.Text)
    mod1.cmd.Parameters("@mm11") = Val(lblMF.Caption) '����MFϵ��,ȷ��ǩ������
    mod1.cmd.Parameters("@mm12") = 0
    mod1.cmd.Parameters("@mm13") = 0
    mod1.cmd.Parameters("@mm14") = 0
    mod1.cmd.Parameters("@mm15") = 0
    mod1.cmd.Parameters("@mm16") = Val(lblMF.Caption)
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    If OptT1.Value = True Then
        mod1.cmd.Parameters("@mb1") = 1 'ͬ��
    Else
        mod1.cmd.Parameters("@mb1") = 0 '�ܾ�
    End If
'''''    If Lc = 1 Then
'''''        mod1.cmd.Parameters("@mb2") = optYj.Value '���������
'''''    Else
'''''        If Val(txtYj.Text) > 0 Then
'''''            mod1.cmd.Parameters("@mb2") = 1 '���������
'''''        End If
'''''    End If
    mod1.cmd.Parameters("@mb2") = YJF '���������
    If txtQy.Text <> "�Ϻ�" And Me.HTLX = "ѯ��ָ��" And (mod1.DName = "������" Or mod1.DName = "�Ǽ���") Then
        mod1.cmd.Parameters("@mb3") = 1
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
        cmdDing.Enabled = False
    End If

    
Set mod1.cmd = Nothing

If mod1.DName = "������" Then
    MsgBox "Hello Xcode!"
End If
End Sub

Private Sub cmdDZ_Click()
Dim ii As Integer
If cmdSave.Enabled = True Then
    MsgBox "���ȱ��棡"
    Exit Sub
End If
'''    Call HTInput(0)
'''    Exit Sub
If Val(cmdDZ.ToolTipText) = 0 Then
    txtHtbh.ToolTipText = cmdDZ.ToolTipText
    Call HTInput(0)
    Exit Sub
End If
If Lc = 1 And LCUid = mod1.DHid Then
    ii = MsgBox("�Ƿ����µ�����Ӻ�ͬ��", vbQuestion + vbYesNo, "����")
    If ii = vbYes Then
        txtHtbh.ToolTipText = cmdDZ.ToolTipText
        Call HTInput(0)
        Exit Sub
    End If
End If
If mod1.DName <> txtYwy.Text And mod1.DName <> txtXYwy.Text And mod1.KhK = 0 And mod1.DName <> "�Ǽ���" And mod1.DName <> "���ĳ�" And mod1.DName <> "����" And mod1.Bm <> "����" And mod1.DName <> "����ϼ" And mod1.DName <> "������" And mod1.DName <> "����" Then Exit Sub

Dim bt() As Byte
Dim tt As String
On Error Resume Next
Kill "c:\work\*.xls": Kill "c:\work\*.doc": Kill "c:\work\*.pdf"
tt = "select fnr,fsize,fname from ht where fid=" & Val(cmdDZ.ToolTipText) & " and xz=0"
frmGGL.adoFile.Recordset.Close
frmGGL.adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockReadOnly, adCmdText

ReDim bt(frmGGL.adoFile.Recordset.Fields("Fsize").Value) As Byte
bt() = frmGGL.adoFile.Recordset.Fields("FNR").GetChunk(frmGGL.adoFile.Recordset.Fields("Fsize").Value + 1)

Open ("c:\work\" & frmGGL.adoFile.Recordset.Fields("fname").Value) For Binary As #2
Put #2, , bt()
Close #2

If Right(frmGGL.adoFile.Recordset.Fields("fname").Value, 3) = "pdf" Then
    MsgBox ("����c:\work\�´򿪴��ļ�")
Else


    frmGGL.OLE2.SourceDoc = "c:\work\" & frmGGL.adoFile.Recordset.Fields("fname").Value
    frmGGL.OLE2.Action = 1
    frmGGL.OLE2.DoVerb (-2)
End If
End Sub

Private Sub cmdDz1_Click()
Dim ii As Integer
If cmdSave.Enabled = True Then
    MsgBox "���ȱ��棡"
    Exit Sub
End If
If Val(cmdDz1.ToolTipText) = 0 Then
    txtHtbh.ToolTipText = cmdDz1.ToolTipText
    Call HTInput(1)
    Exit Sub
End If

If Lc = 1 And LCUid = mod1.DHid Then
    ii = MsgBox("�Ƿ����µ��뼼��������", vbQuestion + vbYesNo, "����")
    If ii = vbYes Then
        txtHtbh.ToolTipText = cmdDz1.ToolTipText
        Call HTInput(1)
        Exit Sub
    End If
End If
If mod1.DName <> txtYwy.Text And mod1.DName <> txtXYwy.Text And mod1.KhK = 0 And mod1.DName <> "�Ǽ���" And mod1.DName <> "���ĳ�" And mod1.DName <> "����" And mod1.Bm <> "����" And mod1.DName <> "����" And mod1.DName <> "������" And mod1.DName <> "����" Then Exit Sub

Dim bt() As Byte
Dim tt As String
On Error Resume Next
Kill "c:\work\*.xls": Kill "c:\work\*.doc"
tt = "select fnr,fsize,fname from ht where fid=" & Val(cmdDz1.ToolTipText) & " and xz=1"
frmGGL.adoFile.Recordset.Close
frmGGL.adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockReadOnly, adCmdText
ReDim bt(frmGGL.adoFile.Recordset.Fields("Fsize").Value) As Byte
bt() = frmGGL.adoFile.Recordset.Fields("FNR").GetChunk(frmGGL.adoFile.Recordset.Fields("Fsize").Value + 1)

Open ("c:\work\" & frmGGL.adoFile.Recordset.Fields("fname").Value) For Binary As #2
Put #2, , bt()
Close #2

If Right(frmGGL.adoFile.Recordset.Fields("fname").Value, 3) = "pdf" Then
    MsgBox ("����c:\work\�´򿪴��ļ�")
Else


    frmGGL.OLE2.SourceDoc = "c:\work\" & frmGGL.adoFile.Recordset.Fields("fname").Value
    frmGGL.OLE2.Action = 1
    frmGGL.OLE2.DoVerb (-2)
End If
End Sub


Private Sub cmdGx_Click()
timZm = 1 '����༭
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
    mod1.cmd.Parameters("@NBLX") = "����༭"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "����"
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtYje.Text)
    mod1.cmd.Parameters("@mm20") = Val(lblFid.Caption)
    mod1.cmd.Parameters("@mb1") = Null
    mod1.cmd.Parameters("@md1") = txtYRQ.Text
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


Private Sub cmdHt_Click()
Dim Ra, Rb, RC, RD, RE, Rf
Dim La, Lb, Lc, Ld, Le, LF
Dim R1, R2, R3
Dim FR As String  '��ͬ������ַ�������ͬ�Ĺ�˾
Dim Qy As String
Dim xZ As String
Dim XZDm As String
Dim tt As String
Dim ii As Integer
Dim oo As Integer
Dim MinRow As Integer
Dim MinStr As String
Dim LX As String
Dim Bid As Long
On Error Resume Next

ii = MsgBox("��ȷ��ǩԼ��˾��" & txtCompanyId.Text & "?", vbYesNo + vbQuestion, "����ע�⣡")
If ii = vbNo Then Exit Sub

dtgNewLx.Col = 4
For oo = 1 To 1000
    dtgNewLx.Row = oo
    If dtgNewLx.Text = "" Then
        Exit For
    End If
    If Trim(dtgNewLx.Text) <> "�����" Then
        MsgBox "����δ�����ɱ�ȷ�ϵ�ѯ�۵�����򿪸�ѯ�۵�ȷ�ϻ�׼�ۣ���ɾ����ѯ�۵���"
        Exit Sub
    End If
    
Next

'dtgNewLx.Row = 1: dtgNewLx.Col = 8: MinRow = Val(dtgNewLx.Text): dtgNewLx.Col = 0: MinStr = dtgNewLx.Text
'ȷ����ͬ���ʺͱ��
MinRow = 0

For oo = 1 To dtgNewLx.Rows
    dtgNewLx.Row = oo
    dtgNewLx.Col = 0
    dtgNewLx.Col = 0
    If dtgNewLx.Text = "ѯ��ָ��" Then
    LX = dtgNewLx.Text
    End If
    If dtgNewLx.Text = "" Then
        Exit For
    End If
    If Not (InStr(1, dtgNewLx.Text, "Ԥ��") > 0) Then
        dtgNewLx.Col = 8
        dtgNewLx.Row = oo

        If MinRow < Val(dtgNewLx.Text) Then
            MinRow = Val(dtgNewLx.Text)
            dtgNewLx.Col = 0: MinStr = dtgNewLx.Text
        End If
    End If
Next
If MinRow = 0 And LX <> "ѯ��ָ��" Then Exit Sub
tt = "select la,lf from newlx where zid=" & MinRow
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
xZ = Ra(0, 0)
XZDm = Ra(1, 0)
If xZ = "ҵ������" Then '��ѯ��ָ��ȡ����
    dtgNewLx.Row = 1: dtgNewLx.Col = 3
    Bid = Right(dtgNewLx.Text, 5)
    tt = "select ywlx from xunjiamx where bid=" & Bid & " and ywlx like '%�˹�%' and delf=1;" & _
     "select ywlx from xunjiamx where bid=" & Bid & " and ywlx like '%����%' and delf=1"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    'Exit Sub
    R1 = mod1.HTP.GetRows
    Set mod1.HTP = mod1.HTP.NextRecordset
    R2 = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    If Not (R1 = Null) Then
        xZ = "�˹�": XZDm = "RG"
    ElseIf Not (R2 = Null) Then
        xZ = "����": XZDm = "CL"
    Else
        xZ = "����": XZDm = "QT"
    End If
    
End If
'Exit Sub
'�ɰ汾2012
'''''''''''�жϺ�ͬ�еĸ�ѯ�۵�����ҵ��Աȷ��
'''''''''''dtgLx.Col = 4
'''''''''''dtgLx.Row = 1
'''''''''''If dtgLx.Text <> "" And W1 = 0 Then
'''''''''''    MsgBox "ά��ѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If
''''''''''' dtgLx.Row = 2
'''''''''''If dtgLx.Text <> "" And W2 = 0 Then
'''''''''''    MsgBox "����ѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If
'''''''''''dtgLx.Row = 3
'''''''''''If dtgLx.Text <> "" And W3 = 0 Then
'''''''''''    MsgBox "�����˹�ѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If
'''''''''''dtgLx.Row = 4
'''''''''''If dtgLx.Text <> "" And W4 = 0 Then
'''''''''''    MsgBox "ѹ����ά�ޱ���ѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If
'''''''''''dtgLx.Row = 5
'''''''''''If dtgLx.Text <> "" And W5 = 0 Then
'''''''''''    MsgBox "ѹ����ó��ѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If
'''''''''''dtgLx.Row = 6
'''''''''''If dtgLx.Text <> "" And W6 = 0 Then
'''''''''''    MsgBox "�н�ҵ��ѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If
'''''''''''dtgLx.Row = 7
'''''''''''If dtgLx.Text <> "" And W7 = 0 Then
'''''''''''    MsgBox "����ѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If
'''''''''''dtgLx.Row = 8
'''''''''''If dtgLx.Text <> "" And W8 = 0 Then
'''''''''''    MsgBox "����ѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If
'''''''''''dtgLx.Row = 9
'''''''''''If dtgLx.Text <> "" And W9 = 0 Then
'''''''''''    MsgBox "�ڴ︻ѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If
'''''''''''dtgLx.Row = 10
'''''''''''If dtgLx.Text <> "" And W10 = 0 Then
'''''''''''    MsgBox "��ͼѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If
'''''''''''dtgLx.Row = 11
'''''''''''If dtgLx.Text <> "" And W11 = 0 Then
'''''''''''    MsgBox "�����ѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If
'''''''''''dtgLx.Row = 12
'''''''''''If dtgLx.Text <> "" And W12 = 0 Then
'''''''''''    MsgBox "�ְ�ѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If
'''''''''''dtgLx.Row = 13
'''''''''''If dtgLx.Text <> "" And W13 = 0 Then
'''''''''''    MsgBox "�Ǵ����Ʒѯ�۵�û�гɱ�ȷ��!"
'''''''''''    Exit Sub
'''''''''''End If



'''''''''''''�жϺ�ͬ���ʺͺ�ͬ���.
''''''''''''If W1 > 0 Or W2 > 0 Or W3 > 0 Then
'''''''''''''''''    ii = MsgBox("��ȷ�ϴ˵�������ǩ������ǩ��" & Chr(13) & Chr(10) & "�������ǡ�������ǩ�������񡱴�����ǩ��", vbYesNo + vbInformation, "����ȷ�ϣ�")
''''''''''''    xZ = "�˹���"
''''''''''''    XZDm = "RG"
''''''''''''ElseIf W4 > 0 Or W5 > 0 Then
''''''''''''    xZ = "ѹ����"
''''''''''''    XZDm = "YS"
''''''''''''ElseIf W6 > 0 Then
''''''''''''    xZ = "�н�"
''''''''''''    XZDm = "ZJ"
''''''''''''ElseIf W7 > 0 Or W8 > 0 Or W9 > 0 Or W10 > 0 Or W11 > 0 Or W12 > 0 Or W13 > 0 Then
''''''''''''    xZ = "ó��"
''''''''''''    XZDm = "TR"
''''''''''''Else
''''''''''''    MsgBox "ֻ��ȷ���˿ͻ��������,�������ɺ�ͬ���!"
''''''''''''    Exit Sub
''''''''''''End If

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
    If txtCompanyId.Text = "�Ϻ���������յ��������޹�˾" Then
        FR = "H"
    ElseIf txtCompanyId.Text = "�Ϻ���������յ��豸���޹�˾" Then
        FR = "D"
    ElseIf txtCompanyId.Text = "�Ϻ�������ó���޹�˾" Then
        FR = "J"
    ElseIf txtCompanyId.Text = "���ݽ�ʨ�����豸���޹�˾" Then
        FR = "S"
    End If
    txtHtbh.Text = FR & "M" & Qy & Format(mod1.DQda, "yyyymmdd") & XZDm & lblHid.Caption
Else
    txtHtbh.Text = "HB" & Qy & Format(mod1.DQda, "yyyymmdd") & XZDm & lblHid.Caption
End If
    lblHtxz.Caption = xZ
    
    '�ɰ汾2012
'''''''''''''    If W1 > 0 Then '��ͬ���ע����ǩ������ǩ
    If MinRow = 2 Then '��ͬ���ע����ǩ������ǩ
        ii = MsgBox("��ȷ�ϴ˵�������ǩ������ǩ��" & Chr(13) & Chr(10) & "�������ǡ�������ǩ�������񡱴�����ǩ��", vbYesNo + vbInformation, "����ȷ�ϣ�")
        If mod1.ZT = "HMData" Then
            If ii = vbYes Then
                txtHtbh.Text = FR & "N" & Qy & Format(mod1.DQda, "yyyymmdd") & XZDm & lblHid.Caption
            Else
                txtHtbh.Text = FR & "O" & Qy & Format(mod1.DQda, "yyyymmdd") & XZDm & lblHid.Caption
            End If
        Else
            If ii = vbYes Then
                txtHtbh.Text = "HN" & Qy & Format(mod1.DQda, "yyyymmdd") & XZDm & lblHid.Caption
            Else
                txtHtbh.Text = "HO" & Qy & Format(mod1.DQda, "yyyymmdd") & XZDm & lblHid.Caption
            End If
        End If
    End If
txtHtxz.Text = xZ
timZm = 11 '���ɺ�ͬ���
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
    mod1.cmd.Parameters("@NBLX") = "��ͬ���"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtHtbh.Text
    mod1.cmd.Parameters("@mt2") = xZ
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = 0
    mod1.cmd.Parameters("@mb1") = 0
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
cmdSave.Enabled = True
End Sub

Private Sub cmdKQy_Click()
FmxcFK.Show
FmxcFK.ZOrder 0
End Sub

Private Sub cmdMod_Click()

If Lc = 1 And LCUid = mod1.DHid Then
    frmFk.Visible = True
    Me.comFPLX.Visible = True
    dt3.Visible = True
    dt4.Visible = True
    companyId.Visible = True
    cmdSave.Enabled = True
    cmdDel.Enabled = True
    frmTj.Visible = True
    '''''optAA.Value = True
    Me.companyId.Visible = True
    txtXYwy.Locked = False
End If
If lblHTF.Caption = "ִ����" Then '��ִͬ�к�ֻ�����ɱ������
    frmTj.Visible = True
    
    'optAb.Value = True
End If
If mod1.Kyj = True And LCUid = mod1.DHid Then
    frmYj.Visible = True
    txtYJ.Locked = False
    cmdSave.Enabled = True
Else
    frmYj.Visible = False
End If
'If (mod1.DName = "�Ǽ���" Or mod1.DName = "������" Or mod1.DName = "����") And lblHTF.Caption = "ִ����" Then
If mod1.DName = "�Ǽ���" Or mod1.DName = "������" Or mod1.DName = "����" Then
    cmdDel.Enabled = True
End If
If mod1.DName = "������" Then
    cmdDel.Enabled = True
    frmFk.Visible = True
    Me.comFPLX.Visible = True
    dt3.Visible = True
    dt4.Visible = True
    companyId.Visible = True
    cmdSave.Enabled = True
    txtYjBz.Locked = False
End If
If mod1.DName = "�޳�" Then
    txtZBZ.Locked = False
    cmdSave.Enabled = True
    Exit Sub
End If
End Sub

Private Sub cmdNQ_Click()
Dim tt As String
Dim Ra
Dim oo As Integer

Dim ii As Integer


On Error Resume Next
optG1.Value = False: optG2.Value = False
txtQM.Text = ""
If (mod1.DName = txtYwy.Text Or mod1.DName = txtXYwy.Text) And Lc > 1 Then
    OptT1.Value = False
Else
    OptT1.Value = True

End If
If txtQy.Text = "�Ϻ�" And mod1.DName <> "�߶���" Then
    Frame1.Visible = False
Else
    Frame1.Visible = True
End If
If Not (lblHTF.Caption = "�༭" Or lblHTF.Caption = "����" Or lblHTF.Caption = "����" Or lblHTF.Caption = "��ִ��") And mod1.DName <> "������" And txtQy.Text = "�Ϻ�" Then
    Exit Sub
End If
If (mod1.DName = "�߶���" Or mod1.DName = "�Ǽ���") And (Me.lblHTF = "ִ����" Or Me.lblHTF = "����" Or Me.lblHTF.Caption = "��ִ��") Then
    LCRen = mod1.DName: LCUid = mod1.DHid

End If

If (mod1.DName = "�Ǽ���" Or mod1.DName = "����" Or mod1.DName = "������") And (lblHTF.Caption = "ִ����" Or lblHTF.Caption = "�༭") Then
    frmQm.Visible = True
    cmdDing.Enabled = True
        optT2.Enabled = True
        OptT1.Value = False
        optT2.Value = False
        Exit Sub
End If
'''''''''''If Lc = 100 Then
'''''''''''    Exit Sub
'''''''''''End If
If LCUid <> mod1.DHid And Not (mod1.DName = txtYwy.Text Or mod1.DName = txtXYwy.Text) Then
    MsgBox "�˴�Ӧ��" & LCRen & "ǩ��! ������Ҫ�ٵ�"
    Exit Sub
End If
'''''''''''If Lc = 100 Then
'''''''''''
'''''''''''        Exit Sub
'''''''''''
'''''''''''End If
If cmdSave.Enabled = True Then
    MsgBox "���Ƚ����ӱ���,��ǩ�����Ĵ���!"
    Exit Sub
End If

    frmQm.Visible = True
    cmdDing.Enabled = True
    
    If Lc = 1 Then   '������ֻ��ǩ�֣����ܲ��ء�
        optT2.Enabled = False
        OptT1.Value = True
    Else
        optT2.Enabled = True
        OptT1.Value = False
        optT2.Value = False
    End If
'''''tt = "select dkz from htping where hid=" & Hid
'''''Set mod1.HTP = CreateObject("adodb.recordset")
'''''mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''Ra = mod1.HTP.GetRows
'''''If mod1.HTP.BOF = True Then
'''''    Set mod1.HTP = Nothing
'''''    Exit Sub
'''''End If
'''''Ra = mod1.HTP.GetRows
'''''mod1.HTP.Close
'''''Set mod1.HTP = Nothing
'''''If Ra(0, 0) = 1 Then
''''''''''    Me.cmd
'''''Else
'''''End If
End Sub

Private Sub cmdNQ_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And mod1.DName = "�Ǽ���" Then
    frmQm.Visible = True
    OptT1.Enabled = False
    optT2.Value = True
End If
End Sub

Private Sub cmdSave_Click()

Dim FPLX As String
Dim tt As String
Dim XYwy As String
Dim XUid As String
Dim Ra
If Me.JCYG = True Then
    MsgBox ("��ͬ��ֻ����һ��Ԥ���ɱ�ѯ�۵���")
    Exit Sub
End If
txtQb.Text = YGCB
If Val(FmxcFK.txtBL1.Text) + Val(FmxcFK.txtBL2.Text) + Val(FmxcFK.txtBL3.Text) <> 100 Then
    MsgBox "û����ȷ���ÿ�����ɱ���!"
    FmxcFK.Show
    FmxcFK.ZOrder 0
    Exit Sub
    
End If
'�Զ�����
If Val(lblHid.Caption) < 26934 Then
    Call Cale
Else
    Call NewCale
End If
FPLX = txtFPLx.Text

'��⿪�ؾ���Ϸ���
tt = "select userid from worker where username='" & txtXYwy.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.BOF = True Then
    MsgBox "���ؾ�����ȷ,��ȷ��!"
    txtXYwy.SetFocus
    Exit Sub
End If
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
XUid = Ra(0, 0)
XYwy = txtXYwy.Text

timZm = 2 '�����ͬ
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
    mod1.cmd.Parameters("@NBLX") = "����"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = FPLX '��Ʊ����
    mod1.cmd.Parameters("@mt25") = FmxcFK.comQy2 '��������
    mod1.cmd.Parameters("@mt26") = FmxcFK.comQy3.Text '��������
    mod1.cmd.Parameters("@mt27") = FmxcFK.txtRen2.Text '��������
    mod1.cmd.Parameters("@mt28") = FmxcFK.txtRen3.Text  '��������
    mod1.cmd.Parameters("@mt29") = FmxcFK.txtRen2.ToolTipText  '��������
    mod1.cmd.Parameters("@mt30") = FmxcFK.txtRen3.ToolTipText  '��������
    mod1.cmd.Parameters("@mt31") = FmxcFK.txtBL1.Text '��������
    mod1.cmd.Parameters("@mt32") = FmxcFK.txtBL2.Text   '��������
    mod1.cmd.Parameters("@mt33") = FmxcFK.txtBL3.Text  '��������
    mod1.cmd.Parameters("@mt34") = txtYjBz.Text '����ע
    mod1.cmd.Parameters("@mt35") = txtZBZ.Text '�޳���ע
    mod1.cmd.Parameters("@mt36") = comQBF.Text  'ȫ����
    mod1.cmd.Parameters("@mt38") = XYwy '���ؾ���
    mod1.cmd.Parameters("@mt39") = XUid
    mod1.cmd.Parameters("@mlt1") = txtBz.Text '��ע
    mod1.cmd.Parameters("@mm1") = D1
    mod1.cmd.Parameters("@mm2") = D2
    mod1.cmd.Parameters("@mm3") = D3
    mod1.cmd.Parameters("@mm4") = D4
    mod1.cmd.Parameters("@mm5") = D5
    mod1.cmd.Parameters("@mm6") = D6
    mod1.cmd.Parameters("@mm7") = D7
    mod1.cmd.Parameters("@mm8") = D8
    mod1.cmd.Parameters("@mm9") = D9
    mod1.cmd.Parameters("@mm10") = D10
    mod1.cmd.Parameters("@mm11") = D11
    mod1.cmd.Parameters("@mm12") = D12
    mod1.cmd.Parameters("@mm13") = D13
    mod1.cmd.Parameters("@mm14") = Val(txtQb.Text)
    'mod1.cmd.Parameters("@mm15") = QBZE
    mod1.cmd.Parameters("@mm21") = Val(txtYJ.Text)
    If txtCompanyId.Text = "�Ϻ���������յ��������޹�˾" Then
        mod1.cmd.Parameters("@mm22") = 1
    ElseIf txtCompanyId.Text = "�Ϻ���������յ��豸���޹�˾" Then
        mod1.cmd.Parameters("@mm22") = 2
    ElseIf txtCompanyId.Text = "�Ϻ�������ó���޹�˾" Then
        mod1.cmd.Parameters("@mm22") = 3
    ElseIf txtCompanyId.Text = "���ݽ�ʨ�����豸���޹�˾" Then
        mod1.cmd.Parameters("@mm22") = 4
    End If
    If chkFP.Value = 1 Then
        mod1.cmd.Parameters("@mb1") = 1
    Else
        mod1.cmd.Parameters("@mb1") = 0
    End If
    If txtF.Text = "" Or txtL.Text = "" Then
    mod1.cmd.Parameters("@md1") = Null
    mod1.cmd.Parameters("@md2") = Null
    Else
    mod1.cmd.Parameters("@md1") = txtF.Text  'ά����ʼ��
    mod1.cmd.Parameters("@md2") = txtL.Text
    End If
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


        
    End If

    
Set mod1.cmd = Nothing
Me.companyId.Visible = False
End Sub

Private Sub cmdTj_Click()
Dim NTJ As Boolean '�ܷ����׷�ӵ�(�ж��б����Ƿ���Ԥ��ѯ�۵�)
Dim QBF As Boolean 'ȫ����
Dim FBF As Boolean '�Ƿ�Ϊ�˹�(�ְ�)�������
Dim Fl As String '׷�ӵ������Ƿ��յ�
Dim ii As Integer
Dim oo As Integer
Dim Ra
Dim Lje As Double
dtgNewLx.Col = 0: dtgNewLx.Row = 1
'''''''If dtgNewLx.Text = "ѯ��ָ��" Then
''''''If dtgNewLx.Text <> "" Then
''''''    'Exit Sub
''''''End If
If lblHTF.Caption = "�༭" Then

Dim tt As String

'''''If mod1.Qy = "�Ϻ�" Then
'''''    FmxcLxNew.cmdNew.Caption = "����ѯ�۵�"
'''''
'''''    FmxcLxNew.Hid = Val(lblHid.Caption)
'''''    FMXCXmmc.txtXMMC.Text = Me.txtXMMC.Text
'''''    FMXCXmmc.txtXMMC.ToolTipText = Me.txtXMMC.ToolTipText
'''''    FMXCXmmc.comKhmc.Text = Me.txtKhmc.Text
'''''    FMXCXmmc.comKhmc.ToolTipText = Me.txtKhmc.ToolTipText
'''''    FmxcLxNew.Show
'''''    FmxcLxNew.ZOrder 0
'''''    Exit Sub
'''''End If

     timZm = 20
     Set mod1.cmd = CreateObject("adodb.command")
     mod1.cmd.ActiveConnection = mod1.workKK
     mod1.cmd.CommandText = "MLAdd"
     mod1.cmd.CommandType = adCmdStoredProc
     mod1.cmd.Parameters("@zid") = 0
     mod1.cmd.Parameters("@errch") = ""
     mod1.cmd.Parameters("@NB") = "�º�ͬ2013"
     mod1.cmd.Parameters("@NBLX") = "���ѯ�۵�"
     mod1.cmd.Parameters("@bh") = ""
     mod1.cmd.Parameters("@ywy") = mod1.DName
     mod1.cmd.Parameters("@uid") = mod1.DHid
     mod1.cmd.Parameters("@mt1") = txtXmmc.Text
     mod1.cmd.Parameters("@mt2") = "ѯ��ָ��"
      mod1.cmd.Parameters("@mt3") = ""
          mod1.cmd.Parameters("@mt4") = lblHid.Caption
     mod1.cmd.Parameters("@mt5") = ""
     mod1.cmd.Parameters("@mt25") = lblHid.Caption
     mod1.cmd.Parameters("@mlt1") = ""
     mod1.cmd.Parameters("@mm1") = Val(txtXmmc.ToolTipText)
     mod1.cmd.Parameters("@mm2") = 0
    ' Exit Sub
     mod1.cmd.Parameters("@md1") = Null
     Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
     mod1.cmd.Execute
    ' MsgBox "b"
     mod1.Zid = mod1.cmd.Parameters("@zid").Value
     If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
         MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
         If timZm = 1 Then
             cmdNew.Enabled = False
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

    

ElseIf lblHTF.Caption = "ִ����" Or lblHTF.Caption = "���" Then

    NTJ = False
    'ȫ����������
    dtgNewLx.Col = 0
    For oo = 1 To dtgNewLx.Rows
        On Error Resume Next
        dtgNewLx.Row = oo

        'If dtgNewLx.Text = "" Then Exit For
        If dtgNewLx.Text = "ȫ������(Ԥ��)" Then
            QBF = True: NTJ = True: Exit For
        ElseIf dtgNewLx.Text = "��ͬ������(Ԥ��)" Then
            QBF = False: NTJ = True: Exit For
        End If
        
    Next
   ' If NTJ = False And Val(txtQb.Text) = 0 Then Exit Sub
    
    If lblHTF.Caption <> "ִ����" Then
        MsgBox "�˺�ͬ������" & lblHTF.Caption & "�׶Σ����������ҵ��!"
        Exit Sub
    End If
    
    If mod1.DName <> "������" And mod1.DName = "�Ϻ�" Then
        Exit Sub
    End If
    
    
    ii = MsgBox("��ȷ�ϴ�׷�ӵ��������(YES)���Ƿְ�(NO)!", vbQuestion + vbYesNo)
    If ii = vbNo Then
        FBF = True
    End If
    
    '����Ƿ񳬳�Ԥ���ɱ�,
    tt = "select sum(ze) from htzuiZe where fl='׷��'and hid=" & Val(lblHid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    If mod1.HTP.BOF = True Then
        MsgBox "����!"
        Exit Sub
    End If
    Ra = mod1.HTP.GetRows
    If IsNull(Ra(0, 0)) = False Then
        Lje = Ra(0, 0)
    End If
    If Lje < Val(txtQb.Text) Then
        Fl = "׷��"
        ii = MsgBox("������׷�ӵ����Ϊ��" & Chr(13) & Chr(10) & "Ԥ���ɱ��ܶ��(" & Val(txtQb.Text) & "�����Ѿ�ʹ�ö�(" & Str(Lje) & "����" & _
        (Val(txtQb.Text) - Lje) & Chr(13) & Chr(10) & "��ȷ��", vbYesNo + vbQuestion, "��ע��")
        If ii = vbNo Then Exit Sub
    Else
        ii = MsgBox("����׷�ӵ����(" & Val(txtQb.Text) & ")��������Ϊ�����ɷ��ջ���", vbYesNo + vbQuestion, "��ע��")
        If ii = vbNo Then Exit Sub
        Fl = "����"
    End If
    
    timZm = 17 '���׷�ӵ�
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�º�ͬ2013"
    mod1.cmd.Parameters("@NBLX") = "���׷�ӵ�"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtHtbh.Text
    mod1.cmd.Parameters("@mt2") = Fl
    If FBF = True Then
        mod1.cmd.Parameters("@mt3") = "�˹�(�ְ�)"
    Else
        mod1.cmd.Parameters("@mt3") = "����"
    End If
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = 0

    mod1.cmd.Parameters("@mb1") = QBF
    mod1.cmd.Parameters("@mb1") = FBF
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


        
    End If
  
    Set mod1.cmd = Nothing

Else

        MsgBox "�˺�ͬ������" & lblHTF.Caption & "�׶Σ����������ҵ��!"

End If
End Sub

Private Sub cmdYadd_Click()
Dim tt As String
Dim Ra
Dim YYY As Long
Dim hg As Single
Dim oo As Integer
On Error Resume Next
dtgLx.Col = 1
For oo = 1 To 12
    dtgYJ.Col = 1
    dtgYJ.Row = oo
    hg = hg + Val(dtgYJ.Text)
Next
tt = "select yj from htping where hid=" & Val(lblHid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
hg = hg + Val(txtYingFu.Text)
If hg > Val(txtYJ.Text) Then
    'Exit Sub
    MsgBox "�������ܳ���Ԥ����!"
    Exit Sub
End If
If (Val(txtFED.Text) = 0 Or Val(txtYingFu.Text) = 0) And mod1.DName <> "������" Then
Exit Sub
End If


'''''''''tt = "select yjff from htping where htbh='" & txtHtbh.Text & "'"
'''''''''Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''If IsNull(mod1.HTP.RecordCount) Or mod1.HTP.RecordCount = 0 Then
'''''''''    MsgBox ("��ȡ���ݴ���1!")
'''''''''    Exit Sub
'''''''''End If
'''''''''If mod1.HTP.Fields("yjff").Value = True Then
'''''''''    MsgBox ("�����Ѿ�ȫ��֧��,�����ٸ���!")
'''''''''    Exit Sub
'''''''''End If


timZm = 16 '��ӽ���
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "����༭"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text) '��ͬ���
    mod1.cmd.Parameters("@mt2") = Trim(txtXmmc.Text) '��Ŀ����
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtFED.Text) / 100 '���
    mod1.cmd.Parameters("@mm2") = Val(txtYingFu.Text) 'Ӧ��
    mod1.cmd.Parameters("@mb1") = 1 '��ӽ���
    mod1.cmd.Parameters("@md1") = Null
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


End Sub

Private Sub cmdYdel_Click()
Dim tt As String
Dim hg As Single
Dim ii As Integer
Dim Yid As Long
Dim Lc As String
On Error Resume Next
dtgYJ.Col = 3
Lc = Val(dtgYJ.Text)
dtgYJ.Col = 2
Yid = Val(dtgYJ.Text)


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

timZm = 16 '����༭
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��ͬ����"
    mod1.cmd.Parameters("@NBLX") = "����༭"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(txtHtbh.Text) '��ͬ���
    mod1.cmd.Parameters("@mt2") = Trim(txtXmmc.Text) '��Ŀ����

    mod1.cmd.Parameters("@mm1") = Yid

    mod1.cmd.Parameters("@mb1") = 0 '����ɾ��

    mod1.cmd.Parameters("@md1") = Null

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


Private Sub cmdYongYou_Click()
Dim tt As String
Dim cSOCode As String
Dim Id As Double
Dim Ra
'1�ȼ�Ȿ��ͬ�Ƿ���ִ��״̬


tt = "select htf from htping where hid=" & Val(lblHid.Caption)

Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workYY, adOpenForwardOnly, adLockReadOnly, adCmdText


'2����������Ƿ��д˵���
3 '���е���ǰ���ݼ��
'ȡ�����۶�����
tt = "select top 1 cSOCode,id from SO_SOMain order by id desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workYY, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.BOF = True Then
    Exit Sub
End If
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
cSOCode = Trim(Str(Ra(0, 0) + 1))
Id = Ra(1, 0) + 1
cSOCode = "0000000007"
tt = "select uid,'������',khmc,khdh from htping where hid=" & Val(lblHid.Caption)

tt = "insert into SO_SOMain (cPersoncode,cmaker,cCusName,cSocode,iExchRate,Ddate,Id,cDefine1,dPreMoDateBT,cDepCode,cBusType,cexch_name,iTaxRate,cSTCode," & _
    "iVTid,dPreDateBT,cCusCode) values" & _
    " ('" & mod1.DHid & "','������','" & txtKhmc.Text & "','" & cSOCode & "',1,getdate()," & Id & ",'" & txtHtbh.Text & "',getdate(),'1','��ͨ����','�����'," & _
    "17,'LP',95,getdate(),'" & txtKhmc.ToolTipText & "');" & _
    "insert into SO_SODetails (cinvCode,iSosId,fSaleCost,ballpurchase,cSOCode,dPreModate,iSum,iNatTax,iQuotedPrice,KL,ID,iNatMoney,iTax,iRowNo,iMoney," & _
    "iNatSum,iTaxUnitPrice,KL2,dPreDate,cInvName,fcusminprice,iUnitPrice,iQuantity) values (" & _
    "'3031F',1000000007,0,0,'" & cSOCode & "',getdate(),50000,50000-50000/1.17,25000,100," & Id & ",21367.52,50000-50000/1.17,1,50000/1.17,50000," & _
    "25000,100,getdate(),'�䶳��',0,5000/2/1.17,2)"
 Set mod1.HTP = CreateObject("adodb.recordset")


mod1.HTP.Open tt, mod1.workYY, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = Nothing
    
''''''''    iSOsID �ӱ�id
''''''''0   fSaleCost ���۵���
''''''''0   ballpurchase �Ƿ�ȫ���ɹ�
''''''''    cSOCode ���۶�����
''''''''    dPreMoDate Ԥ�깤����
''''''''    iSum ԭ�Ҽ�˰�ϼ�
''''''''    iNatTax ����˰��
''''''''    iTaxRate ˰��
''''''''    iQuotedPrice    ���ۣ��Ƿ�˰�ο����ײ���
''''''''100 KL ����
''''''''    iNatUnitPrice ������˰����
''''''''    Id ���۶��������ʶ
''''''''    iNatMoney ������˰���
''''''''    iTax ԭ��˰��
''''''''    iRowNo �к�(�ۼ�ֵ)
''''''''    iMoney ԭ����˰���
''''''''    iNatSum ���Ҽ�˰�ϼ�
''''''''    iTaxUnitPrice ԭ�Һ�˰����
''''''''    KL2 ���ο���
''''''''    dPreDate Ԥ��������
''''''''    cInvName �������
''''''''0   fcusminprice �ͻ�����ۼ�
''''''''    cContractID ��ͬ��(��ͷû��)
''''''''    iUnitPrice ԭ����˰����
''''''''    iQuantity ����
    
    
    
'���ű��ĿǰĬ��Ϊ����ʵ�������ʱ����ȡʵ�ʱ���
'''''''uid cPersonCode ҵ��Ա����
'''''''    cMaker �Ƶ���
'''''''khmc    cCusName    �ͻ�����
'''''''    cSOCode ���۶�����
'''''''1   iExchRate ����
'''''''    Ddate ��������
'''''''    ID ���۶�������
'''''''    cMemo ��ע
'''''''htbh    cDefine1    ������Ϣ��ͬ���
'''''''    dPreMoDateBT Ԥ�깤����
'''''''    cDepCode ���ű���
'''''''    cBusType ҵ������
'''''''    cexch_name ��������
'''''''17  iTaxRate ��ͷ˰��
'''''''    cSTCode �������ͱ���
'''''''95  iVTid ����ģ���
'''''''    dPreDateBT Ԥ��������
'''''''khdh    cCusCode    �ͻ�����
End Sub

Private Sub cmdZX_Click()
FmxcNewZX.Hid = Val(lblHid.Caption)
Call FmxcNewZX.Bound1(Val(lblHid.Caption))
Call FmxcNewZX.Bound2(Val(lblHid.Caption))
Call FmxcNewZX.edQing
FmxcNewZX.Show
FmxcNewZX.ZOrder 0
End Sub

Private Sub comFPLX_Click()
txtFPLx.Text = comFPLX.Text
End Sub


Private Sub comKQY_Click()
txtKQY.Text = comKQY.Text
End Sub


Private Sub companyId_Click()
txtCompanyId.Text = companyId.Text
End Sub

Private Sub dt3_CloseUp()
txtF.Text = dt3.Value
End Sub


Private Sub dt4_CloseUp()
txtL.Text = dt4.Value
End Sub


Private Sub dtgFk_Click()
dtgFKN.Row = dtgFk.Row
dtgFKN.Col = 0: txtYRQ.Text = dtgFKN.Text
dtgFKN.Col = 2: txtYje.Text = dtgFKN.Text
dtgFKN.Col = 3: lblFid.Caption = dtgFKN.Text
End Sub

Private Sub dtgLx_DblClick()
Dim tt As String
Dim ii As Integer
Dim LX As String
Dim Lb As String
Dim Bid As Long
Dim Ra
Dim La As Integer
Dim oo As Integer
On Error Resume Next
FmxcFK.Visible = False
NewId = dtgLx.Row
dtgLx.Col = 0: Lb = dtgLx.Text
dtgLx.Col = 1: LX = dtgLx.Text: XJZL = dtgLx.Text
dtgLx.Col = 4
If dtgLx.Row = 1 Or dtgLx.Row = 2 Or dtgLx.Row = 3 Or dtgLx.Row = 4 Or dtgLx.Row = 6 Or dtgLx.Row = 12 Then
    LLXX = True
Else
    LLXX = False
End If

Bid = Mid(Trim(dtgLx.Text), 4, Len(Trim(dtgLx.Text)) - 3)
If Bid > 0 Then
    mod1.BTZ = 36
    Call FmxcXJ.Bound(Bid)
    dtgLx.Col = 3: Call FmxcXJ.SDJE(Val(dtgLx.Text))
    FmxcXJ.Show
    FmxcXJ.ZOrder 0
    Exit Sub
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
        If dtgLx.Row > 0 And dtgLx.Row < 6 Or dtgLx.Row = 12 Then
            Call frmWBXX.Qing
            Call frmWBXX.Bound(Bid)
            'Call frmWBXNew.Bound(Val(dtglx.Text))
            frmWBXX.Show
            frmWBXX.ZOrder 0
            Exit Sub
        Else
            If mod1.Mname = "������" Or mod1.Mname = "������" Then
                Call frmGxbjNew.Initialize
                frmGxbjNew.Show
                frmGxbjNew.lblTitle.Caption = "<<=��ѡ���ѯ��,����ѡ��ֱ������ԭ�����!"
            Else
                Call modBJD.BJDGXQing

                Call modBJD.BJDBound(Bid, LX)
                If NewId = 7 Then Call frmGXBj.SDJE(D7) '��̯�ٴ���
                If NewId = 8 Then Call frmGXBj.SDJE(D8) '��̯�ٴ���
                If NewId = 9 Then Call frmGXBj.SDJE(D9) '��̯�ٴ���
                If NewId = 10 Then Call frmGXBj.SDJE(D10) '��̯�ٴ���
                If NewId = 11 Then Call frmGXBj.SDJE(D11) '��̯�ٴ���
                If NewId = 13 Then Call frmGXBj.SDJE(D13) '��̯�ٴ���
                Call frmGXBj.dtgMaFF
    
                Call modBJD.gxbjLocked
                frmGXBj.optW.Value = True
                mod1.BTZ = 36
                frmWait.Visible = False
                frmGXBj.Visible = True
                frmGXBj.ZOrder 0
                frmGXBj.cmdMod.Enabled = True
                frmGXBj.cmdSave.Enabled = False
    

            End If
        End If
        Exit Sub
End If

If txtHtbh.Text <> "HMNEW" And mod1.DName <> "������" And mod1.Mname <> "������" Then
    Exit Sub
End If
If Bid = 0 And ((txtYwy.ToolTipText = mod1.DHid Or txtXYwy.ToolTipText = mod1.DHid) And lblHTF.Caption = "�༭" Or mod1.DName = "������") Then

    ii = MsgBox("�Ƿ����ǰ��ѯ�۵��е���" & Lb & "/" & LX & "ѯ�۵�?" & Chr(13) & Chr(10) & "('��'=>���룬'��'=>'�½�')", vbInformation + vbYesNoCancel, "Hello!")
'''    MsgBox ("���ڲ����У�����϶����ã����½⣡")
'''    Exit Sub
    If ii = vbCancel Then
        Exit Sub
    ElseIf ii = vbNo Then
        timZm = 3 '�½�ѯ�۵�
            Set mod1.cmd = CreateObject("adodb.command")
            mod1.cmd.ActiveConnection = mod1.cc
            mod1.cmd.CommandText = "MLAdd"
            mod1.cmd.CommandType = adCmdStoredProc
            mod1.cmd.Parameters("@zid") = 0
            mod1.cmd.Parameters("@errch") = ""
            mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
            mod1.cmd.Parameters("@NBLX") = "�½�ѯ�۵�"
            mod1.cmd.Parameters("@bh") = lblHid.Caption
            mod1.cmd.Parameters("@ywy") = txtXYwy.Text
            mod1.cmd.Parameters("@uid") = txtXYwy.ToolTipText
            mod1.cmd.Parameters("@mt1") = LX
            mod1.cmd.Parameters("@mt2") = txtXmmc.Text
            mod1.cmd.Parameters("@mlt1") = ""
            mod1.cmd.Parameters("@mm1") = 88 'NLBֵ
            mod1.cmd.Parameters("@mm10") = NewId
        
            mod1.cmd.Parameters("@mb1") = LLXX 'LXֵ
            mod1.cmd.Parameters("@md1") = Null
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
    ElseIf ii = vbYes Then
        '''''tt = "select rq,jhg,bid from xunjiaD where delf=1 and xid=" & Val(txtXmmc.ToolTipText) & " and brq> getdate() and (htbh is null or htbh='') and htrow=" & dtgLx.Row & " order by bid desc"
        tt = "select brq,jhg,bid from xunjiaD where delf=1 and xid=" & Val(txtXmmc.ToolTipText) & " and  (htbh is null or htbh='')   order by bid desc"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        La = UBound(Ra, 2) + 1
        Call fmxcXjBr.dtgFF
        For oo = 1 To La + 1
            fmxcXjBr.dtgBr.Row = oo
            fmxcXjBr.dtgBr.Col = 0: fmxcXjBr.dtgBr.Text = Ra(0, oo - 1)
            fmxcXjBr.dtgBr.Col = 1: fmxcXjBr.dtgBr.Text = Ra(1, oo - 1)
            fmxcXjBr.dtgBr.Col = 2: fmxcXjBr.dtgBr.Text = "XJD" & Trim(Str(Ra(2, oo - 1)))
            fmxcXjBr.dtgBr.Col = 3: fmxcXjBr.dtgBr.Text = Ra(2, oo - 1)
            fmxcXjBr.dtgN.Row = oo
            fmxcXjBr.dtgN.Col = 0: fmxcXjBr.dtgN.Text = Ra(0, oo - 1)
            fmxcXjBr.dtgN.Col = 1: fmxcXjBr.dtgN.Text = Ra(1, oo - 1)
            fmxcXjBr.dtgN.Col = 2: fmxcXjBr.dtgN.Text = "XJD" & Trim(Str(Ra(2, oo - 1)))
            fmxcXjBr.dtgN.Col = 3: fmxcXjBr.dtgN.Text = Ra(2, oo - 1)
        Next
        fmxcXjBr.Caption = "ѯ�۵���" & FmxcNew.txtXmmc.Text & "��"
        fmxcXjBr.lblHid.Caption = lblHid.Caption
        fmxcXjBr.Show
        fmxcXjBr.ZOrder 0
    End If
End If
End Sub


Private Sub dtgLx_KeyDown(KeyCode As Integer, Shift As Integer)
MsgBox KeyCode
End Sub

Private Sub dtgLx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ii As Integer
Dim Bid As Long
If Button = 2 And dtgLx.Row > 0 Then
    dtgLx.Col = 4: NewId = dtgLx.Row
   'If txtHtbh.Text <> "HMNEW" Then Exit Sub
    If dtgLx.Text = "" Then Exit Sub
    Bid = Mid(Trim(dtgLx.Text), 4, Len(Trim(dtgLx.Text)) - 3)
    ii = MsgBox("�Ƿ�ȡ����ѯ�۵�?", vbYesNo + vbQuestion, "����")
    If ii = vbNo Then Exit Sub
    timZm = 15 '
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
    mod1.cmd.Parameters("@NBLX") = "ȡ��ѯ�۵�"
    mod1.cmd.Parameters("@bh") = lblHid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtHtbh.Text
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = NewId
    mod1.cmd.Parameters("@mm2") = Bid
    mod1.cmd.Parameters("@mb1") = 0
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
cmdSave.Enabled = True
End If
End Sub


Private Sub dtgNewLx_Click()
'MsgBox dtgNewLx.Row & " " & dtgNewLx.Col
'''''Dim Lrow As Integer
'''''Dim oo As Integer
'''''Lrow = dtgNewLx.Row
'''''On Error Resume Next
''''''����ɫ
'''''For oo = 1 To 21
'''''    dtgNewLx.Row = oo
'''''    dtgNewLx.Col = 1: dtgNewLx.CellForeColor = &H0&
'''''    dtgNewLx.Col = 2: dtgNewLx.CellForeColor = &H0&
'''''    dtgNewLx.Col = 3: dtgNewLx.CellForeColor = &H0&
'''''Next
''''''dtgNewLx.ForeColor = &H0&
'''''
'''''dtgNewLx.Row = Lrow
'''''    dtgNewLx.Col = 1: dtgNewLx.CellForeColor = &HFF&
'''''    dtgNewLx.Col = 2: dtgNewLx.CellForeColor = &HFF&
'''''    dtgNewLx.Col = 3: dtgNewLx.CellForeColor = &HFF&
End Sub

Private Sub dtgNewLx_DblClick()
Dim tt As String
Dim Ld As String
Dim Bid As Long

Dim oo As Integer
On Error Resume Next
dtgNLN.Row = dtgNewLx.Row
dtgNLN.Col = 1
XJZL = dtgNLN.Text

dtgNLN.Col = 5
Ld = dtgNLN.Text
dtgNLN.Col = 6
Bid = Val(dtgNLN.Text)


If Bid > 0 Then
    If Ld = "ѯ�۵�" Then
        mod1.BTZ = 36
        Call FmxcXJ.Bound(Bid)
        dtgNewLx.Col = 2:
        Call FmxcXJ.SDJE(Val(dtgNewLx.Text))
        FmxcXJ.Show
        FmxcXJ.ZOrder 0
    ElseIf Ld = "�ɱ������" Then
        Call fmxcZJ.Bound(Bid)
        fmxcZJ.Show
        fmxcZJ.ZOrder 0
    End If
End If




End Sub


Private Sub dtgNewLx_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Yue As Integer
Dim FM As Single
Dim ii As Integer
On Error Resume Next
If KeyCode = 70 Then
   Yue = DateDiff("D", txtF.Text, txtL.Text)
   Yue = Abs(Yue)
   If Yue = 0 Then Exit Sub
   dtgNewLx.Col = 2
   Yue = Int(Yue / 30)
   FM = Round(Val(dtgNewLx.Text) / Yue, 2)
   ii = MsgBox("������:" & FM, vbInformation + vbOKOnly, Yue & "���·�̯")
   
End If
End Sub

Private Sub dtpYf_CloseUp()
txtYRQ.Text = dtpYf.Value
End Sub


Private Sub Form_Click()
frmFk.Visible = False
frmQm.Visible = False

End Sub
Public Sub dtgPFF()
Dim oo As Integer
For oo = 1 To dtgP.Rows - 1
    dtgP.RowHeight(oo) = dtgP.RowHeight(0)
Next
dtgP.Clear
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "����": dtgP.Col = 1: dtgP.Text = "����": dtgP.Col = 2: dtgP.Text = "ְ��": dtgP.Col = 3: dtgP.Text = "������": dtgP.Col = 4: dtgP.Text = "���":
dtgP.ColWidth(0) = 1005
dtgP.ColWidth(1) = 1005
dtgP.ColWidth(2) = 0
 dtgP.ColWidth(3) = 3000: dtgP.ColWidth(4) = 1035
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub
Public Sub QMBound(Zid As Long, Rz, Lz As Integer)
Dim ii As Integer: Dim oo As Integer
Dim tt As String
On Error Resume Next

dtgP.Rows = Lz + 20
dtgP.Visible = False
Call dtgPFF
For oo = 1 To Lz + 1
    dtgP.Row = oo
    For ii = 0 To 5
        dtgP.Col = ii
        dtgP.Text = Rz(ii, oo - 1)
            DH = 255 * mod1.HH(dtgP.Text, UpInt(dtgP.CellWidth / 100))
            If DH > dtgP.RowHeight(dtgP.Row) Then
                dtgP.RowHeight(dtgP.Row) = DH
            End If
        If ii = 4 Then
            If dtgP.Text = "True" Then
                dtgP.Text = "ͬ��"
            ElseIf dtgP.Text = "False" Then
                dtgP.Text = "����"
            End If

        End If
    Next
Next
For oo = 1 To Lz + 1
    dtgP.Row = oo
    dtgP.Col = 4
            If dtgP.Text = "����" Then
                For ii = 0 To 5
                    dtgP.Col = ii
                    dtgP.CellForeColor = &HFF&
                Next
            End If
Next
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "����": dtgP.Col = 1: dtgP.Text = "����": dtgP.Col = 2: dtgP.Text = "ְ��"
dtgP.Col = 3: dtgP.Text = "������": dtgP.Col = 4: dtgP.Text = "ͨ����"

dtgP.TopRow = 1
dtgP.Visible = True
End Sub

Private Sub Form_DblClick()
If mod1.DName = "������" Or mod1.DName = "�ռ���" Or mod1.DName = "������" Then
    frmYj.Visible = True
End If
If frmYG.Visible = True Then
frmYG.Visible = False
Else
frmYG.Visible = True
txtQb.Locked = True
comQBF.Enabled = False
End If
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth + 500
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0

Call LXFF
Call FKFF
Me.dt3.Value = Date
Me.dt4.Value = Date
Me.dtpYf.Value = Date
FmxcNew.txtBz.Left = 10865
FmxcNew.txtBz.Width = 4150
frmNewLx.Left = 5070
frmNewLx.Top = 0
frmQm.Top = 7320

Call Me.NewLx


End Sub

Public Sub Qing()
Dim oo As Integer
Me.comFPLX.Visible = False
Me.dt3.Visible = False
Me.dt4.Visible = False
cmdHT.Visible = False
Me.companyId.Visible = False
optAA.Value = True
txtHtrq.Text = ""
txtXmmc.Text = "": txtXmmc.ToolTipText = ""
txtKhmc.Text = "": txtKhmc.ToolTipText = ""
txtHtbh.Text = "": txtHtbh.ToolTipText = ""
cmdDZ.ToolTipText = "": cmdDz1.ToolTipText = ""
cmdDZ.Visible = True: cmdDz1.Visible = True
txtHtxz.Text = ""
txtZbh.Text = ""
txtQy.Text = ""
txtBm.Text = ""
txtHtze.Text = ""
txtFPLx.Text = ""
txtZe.Text = ""
txtEd.Text = ""
txtXYwy.Text = "": txtXYwy.ToolTipText = ""
txtYwy.Text = "": txtYwy.ToolTipText = ""
txtF.Text = ""
txtL.Text = ""
lblMF.Caption = ""
lblHTF.Caption = "": lblHTF.ToolTipText = ""
txtYJ.Text = ""
txtYjBz.Text = ""
txtZBZ.Text = ""
Call FKQing
frmFk.Visible = False
txtYRQ.Text = ""
txtYje.Text = ""
lblHid.Caption = ""
txtFX.Text = ""
txtXYwy.Locked = True
optXm.Value = True
optLx.Value = False
txtBz.Text = ""
txtBz.Visible = True
txtQb.Text = ""
txtQB1.Text = ""
chkKDFH.Value = 0 '�����
W1 = 0: W2 = 0: W3 = 0: W4 = 0: W5 = 0: W6 = 0: W7 = 0: W8 = 0: W9 = 0: W10 = 0: W11 = 0: W12 = 0: W13 = 0
D1 = 0: D2 = 0: D3 = 0: D4 = 0: D5 = 0: D6 = 0: D7 = 0: D8 = 0: D9 = 0: D10 = 0: D11 = 0: D12 = 0: D13 = 0
For oo = 1 To 13
    dtgLx.Row = oo
    dtgLx.Col = 2: dtgLx.Text = ""
    dtgLx.Col = 3: dtgLx.Text = ""
    dtgLx.Col = 4: dtgLx.Text = ""
Next
cmdSave.Enabled = False
optYj.Value = Mixed
frmYj.Visible = False
txtCompanyId.Text = "�Ϻ���������յ��������޹�˾"
txtQM.Text = ""

    For ii = 0 To 4
        FmxcNew.Shape1(ii).Visible = False
    Next
txtZBZ.Locked = True
Call dtgYjFF
comQBF.Text = ""
YGCB = 0
QBZE = 0
cmdDing.Enabled = True
frmFP.Visible = False
chkFP.Value = 0
 Me.HTLX = ""

End Sub

Public Sub FKQing()
dtgFk.Clear
dtgFKN.Clear
Call FKFF
End Sub

Public Sub FKFF()
dtgFk.Rows = 60
dtgFk.Cols = 5
dtgFk.Row = 0
dtgFk.Col = 0: dtgFk.Text = "����": dtgFk.CellFontBold = True
dtgFk.Col = 1: dtgFk.Text = "���": dtgFk.CellFontBold = True
dtgFk.Col = 2: dtgFk.Text = "���": dtgFk.CellFontBold = True
dtgFk.Col = 3: dtgFk.Text = "fid": dtgFk.CellFontBold = True
dtgFk.Col = 4: dtgFk.Text = "�����": dtgFk.CellFontBold = True

dtgFk.ColWidth(3) = 0
dtgFk.ColWidth(4) = 0
dtgFk.ColWidth(0) = 1300

dtgFKN.Rows = 60
dtgFKN.Cols = 5
End Sub

Public Sub LXFF()
dtgLx.Row = 0
dtgLx.Col = 0: dtgLx.Text = "ҵ������"
dtgLx.Col = 1: dtgLx.Text = "ҵ������"
dtgLx.Col = 2: dtgLx.Text = "��׼�۸�"
dtgLx.Col = 3: dtgLx.Text = "�ٴ���"
dtgLx.Col = 4: dtgLx.Text = "ѯ�۵�"
dtgLx.Col = 5: dtgLx.Text = "��ͬ���"
dtgLx.Col = 6: dtgLx.Text = "˵��"
dtgLx.MergeCells = flexMergeFree
dtgLx.MergeRow(0) = True
dtgLx.Row = 1: dtgLx.Col = 0: dtgLx.Text = "�˹���"
dtgLx.Row = 1: dtgLx.Col = 1: dtgLx.Text = "ά��": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 1: dtgLx.Col = 6: dtgLx.Text = "����˾��Ա������ɵ��˹�"
dtgLx.Row = 2: dtgLx.Col = 0: dtgLx.Text = "�˹���"
dtgLx.Row = 2: dtgLx.Col = 1: dtgLx.Text = "����": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 2: dtgLx.Col = 6: dtgLx.Text = "����˾��Ա������ɵ��˹�"
dtgLx.Row = 3: dtgLx.Col = 0: dtgLx.Text = "�˹���"
dtgLx.Row = 3: dtgLx.Col = 1: dtgLx.Text = "�����˹�": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 3: dtgLx.Col = 6: dtgLx.Text = "����˾��Ա������ɵ��˹�"
dtgLx.Row = 4: dtgLx.Col = 0: dtgLx.Text = "ѹ����"
dtgLx.Row = 4: dtgLx.Col = 1: dtgLx.Text = "ѹ����ά�ޱ���": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 4: dtgLx.Col = 6: dtgLx.Text = "ѹ����������ά�޻���"
dtgLx.Row = 5: dtgLx.Col = 0: dtgLx.Text = "ѹ����"
dtgLx.Row = 5: dtgLx.Col = 1: dtgLx.Text = "ѹ����ó��": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 5: dtgLx.Col = 6: dtgLx.Text = "ѹ���������Ĳ�Ʒ����"
dtgLx.Row = 6: dtgLx.Col = 0: dtgLx.Text = "�н�"
dtgLx.Row = 6: dtgLx.Col = 1: dtgLx.Text = "�н�ҵ��"
dtgLx.Row = 6: dtgLx.Col = 6: dtgLx.Text = "�н飨�Ӽ䣩ҵ������"
dtgLx.Row = 7: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 7: dtgLx.Col = 1: dtgLx.Text = "����": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 7: dtgLx.Col = 6: dtgLx.Text = "�����豸��ó��"
dtgLx.Row = 8: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 8: dtgLx.Col = 1: dtgLx.Text = "����": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 8: dtgLx.Col = 6: dtgLx.Text = "���ݽ�ʨ���������豸��ó��"
dtgLx.Row = 9: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 9: dtgLx.Col = 1: dtgLx.Text = "�ڴ︻": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 9: dtgLx.Col = 6: dtgLx.Text = "�ڴ︻�豸��ó��"
dtgLx.Row = 10: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 10: dtgLx.Col = 1: dtgLx.Text = "��ͼ": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 10: dtgLx.Col = 6: dtgLx.Text = "��ͼ�豸��ó��"
dtgLx.Row = 11: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 11: dtgLx.Col = 1: dtgLx.Text = "�����": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 11: dtgLx.Col = 6: dtgLx.Text = "����������������׺ģ���ó��"
dtgLx.Row = 12: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 12: dtgLx.Col = 1: dtgLx.Text = "�ְ�"
dtgLx.Row = 12: dtgLx.Col = 6: dtgLx.Text = "�ְ���ͬ"
dtgLx.Row = 13: dtgLx.Col = 0: dtgLx.Text = "ó��"
dtgLx.Row = 13: dtgLx.Col = 1: dtgLx.Text = "�Ǵ����Ʒ"
dtgLx.Row = 13: dtgLx.Col = 6: dtgLx.Text = "�Ǵ����Ʒ��ó��"
dtgLx.Col = 5
dtgLx.Row = 1: dtgLx.Text = "RG": dtgLx.Row = 2: dtgLx.Text = "RG": dtgLx.Row = 3: dtgLx.Text = "RG"
dtgLx.Row = 4: dtgLx.Text = "YS": dtgLx.Row = 5: dtgLx.Text = "YS"
dtgLx.Row = 6: dtgLx.Text = "ZJ"
dtgLx.Row = 7: dtgLx.Text = "TR": dtgLx.Row = 8: dtgLx.Text = "TR": dtgLx.Row = 8: dtgLx.Text = "TR": dtgLx.Row = 10: dtgLx.Text = "TR": dtgLx.Row = 11: dtgLx.Text = "TR"
dtgLx.Row = 12: dtgLx.Text = "TR": dtgLx.Row = 13: dtgLx.Text = "TR": dtgLx.Row = 9: dtgLx.Text = "TR"
dtgLx.MergeCol(5) = True
dtgLx.MergeCol(0) = True
dtgLx.ColWidth(1) = 1695
dtgLx.ColWidth(2) = 1000
dtgLx.ColWidth(3) = 1000
dtgLx.ColWidth(4) = 1000
dtgLx.ColWidth(5) = 1000
dtgLx.ColWidth(6) = 3000
End Sub

Public Sub Bound(Hid As Long)
Dim tt As String
Dim Ra, Rb, RC, Rz, RD, RE, Rf, Rg, Rh
Dim Lz As Integer
Dim Lb As Integer
Dim Ld As Integer
Dim LF As Integer
Dim oo As Integer
Call Qing
Me.Enabled = True
tt = "declare @htbh nvarchar(22),@LcUid nvarchar(22);" & _
    "select @htbh=htbh,@LcUid=lcuid from htping where hid=" & Hid & ";" & _
    "select htrq,xmmc,xid,khmc,khdh,htbh,htxz,zbh,qy,bm,lc,htze,fplx,xywy,xuid,ywy,uid,htqy,htqy1,htf,yj,lcren,lcuid," & _
    "kren,kuid,kqy,kren2,kuid2,kqy2,klb0,klb,klb2,bz,w1,w2,w3,w4,w5,w6,w7,w8,w9,w10,w11,w12,w13,d1,d2,d3,d4,d5,d6,d7,d8,d9,d10,d11,d12,d13, " & _
    "bid1,bid2,bid3,bid4,bid5,bid6,bid7,bid8,bid9,bid10,bid11,bid12,bid13,companyid,fwid,yjbz,clf,addZd1,yy,delf,qbf,newf,qbze,fpf from htping where hid=" & Hid & ";" & _
        "select Ӧ������,�տ���,Ӧ�����,fid,kdfh from htFK where htbh='" & Hid & "';" & _
             "select fid from hmht.dbo.ht where htbh=@htbh and xz=0;" & _
            "select trq,ywy,zn,bz,tf from pizu where bh='" & Hid & "' and yid=80 order by pid desc;" & _
            "select yed,yingFu,yid,lc from yongjin where htbh=@htbh;" & _
            "select fid from hmht.dbo.ht where htbh=@htbh and xz=1;" & _
            "select zl,jhg,sdje,'BJD'+cast(bid as nvarchar(20)),lc,0,bid,lcren,htrow from xunjiaD where htbh='" & Trim(Str(Hid)) & "' and delf=1 order by bid;" & _
            "select dbo.htzui.zl,sum(dbo.htzuidetail.ze) as Ze,0,dbo.htzui.bh,dbo.htzui.lc,1,dbo.htzui.zid,dbo.htzui.lcren,dbo.htzui.htrow" & _
            "  FROM dbo.htZui LEFT OUTER JOIN dbo.htZuiDetail ON dbo.htZui.Zid = dbo.htZuiDetail.Zid where dbo.htzui.hid=" & Hid & " and dbo.htzui.delf=1 " & _
            " group by dbo.htzui.zl,dbo.htzui.bh,dbo.htzui.lc,dbo.htzui.lcren,dbo.htzui.zid,dbo.htzui.htrow order by dbo.htzui.zid;" & _
            "if @lcuid='" & mod1.DHid & "'  update htping set dkz=1 where hid=" & Hid & ";" & _
            "select sum(je) from htAview where htbh=@htbh and lc=100 group by htbh"


            
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.BOF = True Then
    MsgBox "����!"
    Exit Sub
End If
Ra = mod1.HTP.GetRows
On Error Resume Next
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows '��������
Set mod1.HTP = mod1.HTP.NextRecordset
Rz = mod1.HTP.GetRows 'ǩ������
Set mod1.HTP = mod1.HTP.NextRecordset
RD = mod1.HTP.GetRows '����
Set mod1.HTP = mod1.HTP.NextRecordset
RE = mod1.HTP.GetRows '��������
Set mod1.HTP = mod1.HTP.NextRecordset
Rf = mod1.HTP.GetRows '��ҵ����ϸ
Set mod1.HTP = mod1.HTP.NextRecordset
Rg = mod1.HTP.GetRows '�ɱ������
Set mod1.HTP = mod1.HTP.NextRecordset
Rh = mod1.HTP.GetRows 'ʵ���տ�
mod1.HTP.Close

Set mod1.HTP = Nothing
If Rf(0, 0) = "ѯ��ָ��" Then
    Me.HTLX = "ѯ��ָ��"
End If
txtHtrq.Text = Ra(0, 0)
If Year(txtHtrq.Text) >= "2013" Then
    frmYG.Visible = True
Else
    frmYG.Visible = False
End If
txtXmmc.Text = Ra(1, 0): txtXmmc.ToolTipText = Ra(2, 0)
txtKhmc.Text = Ra(3, 0): txtKhmc.ToolTipText = Ra(4, 0)
txtHtbh.Text = Ra(5, 0)
txtHtxz.Text = Ra(6, 0)
txtZbh.Text = Ra(7, 0)
txtQy.Text = Ra(8, 0)
txtBm.Text = Ra(9, 0)

txtHtze.Text = Ra(11, 0)
txtFPLx.Text = Ra(12, 0)
txtXYwy.Text = Ra(13, 0): txtXYwy.ToolTipText = Ra(14, 0)
txtYwy.Text = Ra(15, 0): txtYwy.ToolTipText = Ra(16, 0)
txtF.Text = Ra(17, 0)
txtL.Text = Ra(18, 0)
lblHTF.ToolTipText = Ra(19, 0)

txtYJ.Text = Ra(20, 0)
Lc = Ra(10, 0)
Select Case lblHTF.ToolTipText
Case 0
    lblHTF.Caption = "�༭"
Case 6
    lblHTF.Caption = "����"
Case 9
    lblHTF.Caption = "����"
Case 1
    lblHTF.Caption = "��ִ��"
Case 2
    lblHTF.Caption = "���"
Case 3
    lblHTF.Caption = "ִ����"
Case 100
    lblHTF.Caption = "���"
End Select
If Lc > 1 Then
    optYj.Visible = False
Else
    optYj.Visible = True
End If
If Val(txtYJ.Text) > 0 Then
    optYj.Value = Checked
End If
LCRen = Ra(21, 0)
LCUid = Ra(22, 0)

'��������
FmxcFK.comQy1.Text = txtQy.Text
FmxcFK.txtRen1.Text = txtYwy.Text
FmxcFK.txtRen1.ToolTipText = txtYwy.ToolTipText
FmxcFK.txtBL1.Text = Ra(29, 0)
FmxcFK.comQy2.Text = Ra(25, 0)
FmxcFK.txtRen2.Text = Ra(23, 0)
FmxcFK.txtRen2.ToolTipText = Ra(24, 0)
FmxcFK.txtBL2.Text = Ra(30, 0)
FmxcFK.comQy3.Text = Ra(28, 0)
FmxcFK.txtRen3.Text = Ra(26, 0)
FmxcFK.txtRen3.ToolTipText = Ra(27, 0)
FmxcFK.txtBL3.Text = Ra(31, 0)

FmxcNew.txtBz.Text = Ra(32, 0)

'���������
W1 = Ra(33, 0): W2 = Ra(34, 0): W3 = Ra(35, 0): W4 = Ra(36, 0): W5 = Ra(37, 0): W6 = Ra(38, 0)
W7 = Ra(39, 0): W8 = Ra(40, 0): W9 = Ra(41, 0): W10 = Ra(42, 0): W11 = Ra(43, 0): W12 = Ra(44, 0): W13 = Ra(45, 0)
D1 = Ra(46, 0): D2 = Ra(47, 0): D3 = Ra(48, 0): D4 = Ra(49, 0): D5 = Ra(50, 0): D6 = Ra(51, 0)
D7 = Ra(52, 0): D8 = Ra(53, 0): D9 = Ra(54, 0): D10 = Ra(55, 0): D11 = Ra(56, 0): D12 = Ra(57, 0): D13 = Ra(58, 0)
'''''dtgLx.Row = 1: dtgLx.Col = 2
'''''If W1 > 0 Then
'''''    dtgLx.Text = W1
'''''End If
For oo = 1 To 13
    dtgLx.Row = oo
    Select Case oo
    Case 1
        dtgLx.Col = 2: If W1 > 0 Then dtgLx.Text = W1: dtgLx.Col = 3: If D1 > 0 Then dtgLx.Text = D1
    Case 2
        dtgLx.Col = 2: If W2 > 0 Then dtgLx.Text = W2: dtgLx.Col = 3: If D2 > 0 Then dtgLx.Text = D2
    Case 3
        dtgLx.Col = 2: If W3 > 0 Then dtgLx.Text = W3: dtgLx.Col = 3: If D3 > 0 Then dtgLx.Text = D3
    Case 4
        dtgLx.Col = 2: If W4 > 0 Then dtgLx.Text = W4: dtgLx.Col = 3: If D4 > 0 Then dtgLx.Text = D4
    Case 5
        dtgLx.Col = 2: If W5 > 0 Then dtgLx.Text = W5: dtgLx.Col = 3: If D5 > 0 Then dtgLx.Text = D5
    Case 6
        dtgLx.Col = 2: If W6 > 0 Then dtgLx.Text = W6: dtgLx.Col = 3: If D6 > 0 Then dtgLx.Text = D6
    Case 7
        dtgLx.Col = 2: If W7 > 0 Then dtgLx.Text = W7: dtgLx.Col = 3: If D7 > 0 Then dtgLx.Text = D7
    Case 8
        dtgLx.Col = 2: If W8 > 0 Then dtgLx.Text = W8: dtgLx.Col = 3: If D8 > 0 Then dtgLx.Text = D8
    Case 9
        dtgLx.Col = 2: If W9 > 0 Then dtgLx.Text = W9: dtgLx.Col = 3: If D9 > 0 Then dtgLx.Text = D9
    Case 10
        dtgLx.Col = 2: If W10 > 0 Then dtgLx.Text = W10: dtgLx.Col = 3: If D10 > 0 Then dtgLx.Text = D10
    Case 11
        dtgLx.Col = 2: If W11 > 0 Then dtgLx.Text = W11: dtgLx.Col = 3: If D11 > 0 Then dtgLx.Text = D11
    Case 12
        dtgLx.Col = 2: If W12 > 0 Then dtgLx.Text = W12: dtgLx.Col = 3: If D12 > 0 Then dtgLx.Text = D12
    Case 13
        dtgLx.Col = 2: If W13 > 0 Then dtgLx.Text = W13: dtgLx.Col = 3: If D13 > 0 Then dtgLx.Text = D13
    End Select
    dtgLx.Col = 4: If Ra(58 + oo, 0) > 0 Then dtgLx.Text = "XJD" & Ra(58 + oo, 0)
Next
If Ra(72, 0) = 1 Then
    txtCompanyId.Text = "�Ϻ���������յ��������޹�˾"
ElseIf Ra(72, 0) = 2 Then
    txtCompanyId.Text = "�Ϻ���������յ��豸���޹�˾"
ElseIf Ra(72, 0) = 3 Then
    txtCompanyId.Text = "�Ϻ�������ó���޹�˾"
ElseIf Ra(72, 0) = 4 Then
    txtCompanyId.Text = "���ݽ�ʨ�����豸���޹�˾"
End If
Fwid = Ra(73, 0)
txtYjBz.Text = Ra(74, 0)
txtQb.Text = Ra(75, 0)

txtZBZ.Text = Ra(76, 0)


txtHtbh.ToolTipText = Hid
cmdDZ.ToolTipText = RC(0, 0) '���������fid
cmdDz1.ToolTipText = RE(0, 0) '���������fid

If txtHtbh.Text = "HMNEW" Then Me.cmdHT.Visible = True
txtZe.Text = Rh(0, 0)
txtEd.Text = Round(Val(txtZe.Text) / Val(txtHtze.Text) * 100, 2)
lblMF.Caption = ""
lblHid.Caption = Hid
Lb = UBound(Rb, 2) + 1
Call FKBound(Rb, Lb)
Call Cale '�Զ�����
Lz = UBound(Rz, 2) + 1
Call QMBound(Str(lblHid.Caption), Rz, Lz)
If Lc = 1 Then
    optYj.Value = Mixed
End If
If Lc > 1 Then
optYj.Visible = False
End If
lblTX.Caption = "������:  " & LCRen
If Lc = 100 Then lblTX.Caption = "��ͬ�Ѿ�ִ�У����̽�����"
If Ra(78, 0) = False Then
    lblTX.Caption = "�˵������ϣ�����ԭ��" & Ra(77, 0)
End If
comQBF.Text = Ra(79, 0) 'ȫ����


Ld = UBound(RD, 2) + 1
For oo = 1 To Ld
    dtgYJ.Row = oo
    dtgYJ.Col = 0: dtgYJ.Text = RD(0, oo - 1)
    dtgYJ.Col = 1: dtgYJ.Text = RD(1, oo - 1)
    dtgYJ.Col = 2: dtgYJ.Text = RD(2, oo - 1)
    dtgYJ.Col = 3: dtgYJ.Text = RD(3, oo - 1)
Next
Call Me.LXBound(Rf, Rg)
If Ra(80, 0) = 8 Then 'newF
    frmNewLx.Visible = True
    optLx.Visible = False: optYj.Visible = False
    optXm.Visible = False
Else
    frmNewLx.Visible = False
    optLx.Visible = True: optYj.Visible = True
    optXm.Visible = True
End If

QBZE = Ra(81, 0)
If Ra(82, 0) = True Then
    frmFP.Visible = True
    chkFP.Value = 1
Else
    frmFP.Visible = False
    chkFP.Value = 0
End If
Call NewCale
txtQB1.Text = Round(Val(txtQb.Text) / 2.2, 2)
'''''If lblHTF.Caption = "ִ����" Then '��ִͬ�У�ֻ�����ɱ������
'''''    optAb.Value = True
'''''Else
'''''    optAA.Value = True
'''''End If
optAA.Enabled = True: optAb.Enabled = True
If (mod1.DName = "������" Or mod1.DName = "�Ǽ���") And (Me.lblHTF = "ִ����" Or Me.lblHTF = "����" Or Me.lblHTF.Caption = "��ִ��") And Lc <> 100 Then
    LCRen = mod1.DName: LCUid = mod1.DHid
lblTX.Caption = "������:  " & LCRen
End If
If mod1.GxName = "���۹���" And mod1.GXF = True And Me.HTLX = "ѯ��ָ��" Then
    cmdBJ.Visible = True
Else
    cmdBJ.Visible = False
End If
End Sub

Public Sub FKBound(Rb, Lb As Integer)
Dim FK As Single
Dim oo As Integer
Call FKQing
Call FKFF
On Error Resume Next
For oo = 1 To Lb
    
    dtgFk.Row = oo
    dtgFk.Col = 0: dtgFk.Text = Rb(0, oo - 1): dtgFKN.Col = 0: dtgFKN.Text = Rb(0, oo - 1)
    dtgFk.Col = 2: dtgFk.Text = Rb(2, oo - 1): FK = Rb(2, oo - 1)
    dtgFk.Col = 1: dtgFk.Text = Str(Round(FK / Val(txtHtze.Text), 2) * 100) & "%"
    dtgFk.Col = 3: dtgFk.Text = Rb(3, oo - 1)
    If Rb(4, oo - 1) = True Then
        dtgFk.Col = 4: dtgFk.Text = "��"
        dtgFk.Col = 0: dtgFk.Text = "�����"
        dtgFk.CellAlignment = 0
        dtgFKN.Col = 0: dtgFKN.Text = "�����"
    End If
    dtgFKN.Row = oo

    dtgFKN.Col = 1: dtgFKN.Text = Str(Round(FK / Val(txtHtze.Text), 2) * 100) & "%"
    dtgFKN.Col = 2: dtgFKN.Text = Rb(2, oo - 1)
    dtgFKN.Col = 3: dtgFKN.Text = Rb(3, oo - 1)
    dtgFKN.Col = 4: dtgFKN.Text = dtgFk.Text

Next
End Sub

Private Sub NiceOption1_Click()
txtBz.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If mod1.Kyj = True Then
    If X > (Me.Width - 1000) And Y < 1000 Then
        frmYj.Visible = True
        
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
If htBrow.Visible = True Then
    Call htBrow.dtgREF
    htBrow.Enabled = True
    htBrow.ZOrder 0
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
ElseIf FmxcXB.Visible = True Then
    FmxcXB.Enabled = True
    FmxcXB.ZOrder 0
End If

FmxcFK.Visible = False
Cancel = True
End Sub

Private Sub Label12_DblClick()
Dim tt As String
Dim oo As Integer
Dim Ra
Dim La
If frmNewLx.Visible = True Then
    Exit Sub
End If
If lblHTF.Caption <> "ִ����" And lblHTF.Caption <> "���" And lblHTF.Caption <> "��ִ��" Or Val(txtQb.Text) = 0 Then
    Exit Sub
End If
FmxcZBR.dtgZBr.Clear: FmxcZBR.dtgN.Clear
FmxcZBR.dtgFF
FmxcZBR.Show
FmxcZBR.ZOrder 0
tt = "select bh,gui,ze,zid from htZui where hid=" & Val(lblHid.Caption) & " order by zid"
tt = "SELECT dbo.htZui.Bh, dbo.htZui.Gui, SUM(dbo.htZuiDetail.Ze) AS Ze, dbo.htZui.Zid FROM dbo.htZui LEFT OUTER JOIN dbo.htZuiDetail ON dbo.htZui.Zid = dbo.htZuiDetail.Zid" & _
    " where dbo.htzui.hid=" & lblHid.Caption & " and htzui.delf=1  GROUP BY dbo.htZui.Bh, dbo.htZui.Gui, dbo.htZui.Zid order by dbo.htzui.zid"
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

Private Sub NiceButton2_Click()
If Val(txtHtbh.ToolTipText) = 0 Then
    Call HTInput(1)
    Exit Sub
End If
If mod1.DName <> txtYwy.Text And mod1.DName <> txtXYwy.Text And mod1.KhK = 0 And mod1.DName <> "�Ǽ���" And mod1.DName <> "�Ǽ���" And mod1.DName <> "����" And mod1.Bm <> "����" And mod1.DName <> "����" And mod1.DName <> "�Ǽ���" And mod1.DName <> "����" Then Exit Sub

Dim bt() As Byte
Dim tt As String
On Error Resume Next
Kill "c:\work\*.xls": Kill "c:\work\*.doc"
tt = "select fnr,fsize,fname from ht where fid=" & Val(txtHtbh.ToolTipText)
frmGGL.adoFile.Recordset.Close
frmGGL.adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockReadOnly, adCmdText
ReDim bt(frmGGL.adoFile.Recordset.Fields("Fsize").Value) As Byte
bt() = frmGGL.adoFile.Recordset.Fields("FNR").GetChunk(frmGGL.adoFile.Recordset.Fields("Fsize").Value + 1)

Open ("c:\work\" & frmGGL.adoFile.Recordset.Fields("fname").Value) For Binary As #2
Put #2, , bt()
Close #2

    frmGGL.OLE2.SourceDoc = "c:\work\" & frmGGL.adoFile.Recordset.Fields("fname").Value
    frmGGL.OLE2.Action = 1
    frmGGL.OLE2.DoVerb (-2)
End Sub

Private Sub optAA_Click()
dtgNewLx.Visible = False
Call Me.LXBound(Rf, Rg)
dtgNewLx.Visible = True
Call NewCale
End Sub

Private Sub optAb_Click()
dtgNewLx.Visible = False
Call Me.LXBound1(Rf, Rg)
dtgNewLx.Visible = True
End Sub


Private Sub OptAc_Click()
dtgNewLx.Visible = False
Call Me.LXBound2(Rf, Rg)
dtgNewLx.Visible = True
End Sub

Private Sub optG1_Click()
txtQM.Text = "�Ѹ���"
End Sub

Private Sub optG2_Click()
txtQM.Text = "����ԭ��"
End Sub


Private Sub optLx_Click()
txtBz.Visible = False
End Sub

Private Sub optXm_Click()
txtBz.Visible = True
End Sub

Private Sub optYj_Click()
Dim tt As String
Dim Ra
If optYj.Value = Checked Then
    tt = "select khman,khmob from khren where khdh='" & txtKhmc.ToolTipText & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    dtgRen.Clear
    dtgRen.Row = 0
    
    
End If
End Sub

Private Sub timQuit_Timer()
Dim Rz
Dim Lz As Integer
Dim Rb
Dim Lb As Integer
Dim RD
Dim Ld As Integer
On Error Resume Next
Dim ii As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '���Ϊ����༭
    tt = "select Ӧ������,�տ���,Ӧ�����,fid,kdfh from htFK where htbh='" & lblHid.Caption & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rb = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lb = UBound(Rb, 2) + 1
    Call FmxcNew.FKBound(Rb, Lb)
ElseIf timZm = 2 Then '����
    cmdSave.Enabled = False
    Me.comFPLX.Visible = False
    Me.dt3.Visible = False
    Me.dt4.Visible = False
    frmFk.Visible = False
ElseIf timZm = 10 Then 'ǩ��
    tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & lblHid.Caption & "' and yid=80 order by pid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rz = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lz = UBound(Rz, 2) + 1
    Call Me.QMBound(Val(lblHid.Caption), Rz, Lz)
    frmQm.Visible = False
    If lblHTF.Caption = "ִ����" Then
        'MsgBox "�Ѿ��ɹ�֪ͨ��ִͬ��!"
    End If
ElseIf timZm = 11 Then
    cmdHT.Visible = False
    If W1 > 0 Then
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
ElseIf timZm = 15 Then
    dtgLx.Row = NewId
    dtgLx.Col = 2: dtgLx.Text = ""
    dtgLx.Col = 3: dtgLx.Text = ""
    dtgLx.Col = 4: dtgLx.Text = ""
ElseIf timZm = 16 Then '����༭
    tt = "select yed,yingFu,yid,lc from yongjin where htbh='" & txtHtbh.Text & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    If mod1.HTP.BOF = True Then
        Set mod1.HTP = Nothing
    Else
        RD = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        Ld = UBound(RD, 2) + 1
        Call YjBound(RD, Ld)
    End If
ElseIf timZm = 17 Then '���׷�ӵ�
    optAb.Value = True
    Call Me.LXBound1(Rf, Rg)
ElseIf timZm = 20 Then '��ѯ��ָ��
    Call FmxcXJ.Bound(Me.Bid)
    FmxcXJ.Show
    FmxcXJ.ZOrder 0
    If mod1.Qy = "�Ϻ�" Then
    MsgBox "���ڱ�ע����д��Ҫѯ�۵����ݣ�"
    End If
End If
timQuit.Enabled = False


End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim Bid As Long
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mm2,mt2,mt1,mt3,mt4 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '�ύ�ɹ�
    mod1.Ti = 5
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    timWait.Enabled = False
    If timZm = 1 Then
        txtHtze.Text = mod1.WP.Fields("mm1").Value
    ElseIf timZm = 3 Then
        Bid = mod1.WP.Fields("mt2").Value
        Call FmxcXJ.Bound(Bid)
        FmxcXJ.Show
        FmxcXJ.ZOrder 0
        If NewId = 1 Or NewId = 2 Or NewId = 3 Or NewId = 4 Or NewId = 6 Or NewId = 12 Then
            FmxcXJ.frmWB.Visible = True
        Else
            FmxcXJ.frmSd.Visible = True
        End If
        FmxcXJ.cmdSave.Enabled = True
        FmxcNew.dtgLx.Col = 4
        FmxcNew.dtgLx.Row = NewId: FmxcNew.dtgLx.Text = "XJD" & Trim(Str(Bid))
    ElseIf timZm = 10 Then 'ǩ��
        Lc = mod1.WP.Fields("mm1").Value
        Fwid = mod1.WP.Fields("mm2").Value
        LCRen = mod1.WP.Fields("mt1").Value
        LCUid = mod1.WP.Fields("mt2").Value
        lblTX.Caption = "��һ����,������ " & LCRen
        If Lc = 100 Then lblTX.Caption = "��ͬ�Ѿ�ִ�У����̽�����"
        txtZbh.Text = mod1.WP.Fields("mt3").Value
        lblHTF.ToolTipText = mod1.WP.Fields("mt4").Value
        Select Case lblHTF.ToolTipText
        Case 0
            lblHTF.Caption = "�༭"
        Case 6
            lblHTF.Caption = "����"
        Case 9
            lblHTF.Caption = "����"
        Case 1
            lblHTF.Caption = "��ִ��"
        Case 2
            lblHTF.Caption = "���"
        Case 3
            lblHTF.Caption = "ִ����"
        End Select
    ElseIf timZm = 17 Then '���׷��(����)��
        
    
        Call fmxcZJ.Bound(mod1.WP.Fields("mm1").Value)
        fmxcZJ.Show
        fmxcZJ.ZOrder 0
    ElseIf timZm = 20 Then
        Me.Bid = Val(mod1.WP.Fields("bh").Value)
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then 'δ���

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '����ʧ��
    ii = MsgBox("���������ڴ�����������ʱ,�������´���:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "��������!")
    timWait.Enabled = False
    Unload frmWaitA
    Me.Enabled = True
    Exit Sub
'''''    If timZm = 1 Then
'''''        NiceButton1.Enabled = False
'''''    End If
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("���������ڴ�����������ʱ,��ʱ!", vbExclamation + vbOKOnly, "��������!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
'''''    If timZm = 1 Then
'''''        NiceButton1.Enabled = False
'''''    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub



Public Sub Cale()

dtgLx.Col = 2: dtgLx.Row = 1: W1 = Val(dtgLx.Text)
dtgLx.Col = 2: dtgLx.Row = 2: W2 = Val(dtgLx.Text)
dtgLx.Col = 2: dtgLx.Row = 3: W3 = Val(dtgLx.Text)
dtgLx.Col = 2: dtgLx.Row = 4: W4 = Val(dtgLx.Text)
dtgLx.Col = 2: dtgLx.Row = 5: W5 = Val(dtgLx.Text)
dtgLx.Col = 2: dtgLx.Row = 6: W6 = Val(dtgLx.Text)
dtgLx.Col = 2: dtgLx.Row = 7: W7 = Val(dtgLx.Text)
dtgLx.Col = 2: dtgLx.Row = 8: W8 = Val(dtgLx.Text)
dtgLx.Col = 2: dtgLx.Row = 9: W9 = Val(dtgLx.Text)
dtgLx.Col = 2: dtgLx.Row = 10: W10 = Val(dtgLx.Text)
dtgLx.Col = 2: dtgLx.Row = 11: W11 = Val(dtgLx.Text)
dtgLx.Col = 2: dtgLx.Row = 12: W12 = Val(dtgLx.Text)
dtgLx.Col = 2: dtgLx.Row = 13: W13 = Val(dtgLx.Text)

'Me.lblCBZE.Caption = W1 + W2 + W3 + W4 + W5 + W6 + W7 + W8 + W9 + W10 + W11 + W12 + W13 + Val(txtQb.Text)
Me.lblCBZE.Caption = W1 + W2 + W3 + W4 + W5 + W6 + W7 + W8 + W9 + W10 + W11 + W12 + W13
Me.lblRGF.Caption = W1 + W2 + W3
Me.lblYs.Caption = W4 + W5
Me.lblZJ.Caption = W6
Me.lblMy.Caption = W7 + W8 + W9 + W10 + W11 + W12 + W13
Me.lblLr.Caption = Val(txtHtze.Text) - Val(lblCBZE.Caption)
If Val(lblCBZE.Caption) > 0 Then
Me.lblMF.Caption = Round((Val(txtHtze.Text) - Val(txtYJ.Text) - Val(txtQb.Text)) / Val(lblCBZE.Caption), 2)
End If
Call DJ
End Sub

Public Sub NewCale()
Dim oo As Integer
Dim ZE As Double
Dim LXG As Double
Dim Lje As Double
Dim LYG As Double
Dim htQY As Double 'ȥ��Ԥ���ɱ���ĺ�ͬ�ܽ�������������ٴ���
Dim YGSD As Double 'Ԥ���ٴ���

'����ɱ��ܶ�
ZE = 0
For oo = 1 To Me.dtgNLN.Rows
    dtgNLN.Row = oo
    dtgNLN.Col = 0: If dtgNLN.Text = "" Then Exit For
    dtgNLN.Row = oo
    dtgNLN.Col = 4
    If dtgNLN.Text = "�����" Then
        dtgNLN.Col = 0
        If Not (InStr(1, dtgNLN, "Ԥ��") > 0) Then
            dtgNLN.Col = 1
            ZE = ZE + Val(dtgNLN.Text)
        End If
    End If
Next


Me.lblCBZE.Caption = ZE
'Exit Sub
If Val(lblCBZE.Caption) > 0 Then
Me.lblMF.Caption = Round((Val(txtHtze.Text) - Val(txtYJ.Text) - QBZE) / Val(lblCBZE.Caption), 2)
End If
On Error Resume Next
'�ٴ���
LXG = 0: Lje = 0
htQY = Val(txtHtze.Text)
'ȥ��Ԥ�����
For oo = 1 To Me.dtgNewLx.Rows
    dtgNewLx.Row = oo
    dtgNewLx.Col = 0
    If dtgNewLx.Text = "" Then Exit For

        If InStr(1, dtgNewLx.Text, "Ԥ��") > 0 Then
            'dtgNewLx.Col = 2
            LYG = Val(txtQb.Text)
            dtgNewLx.Col = 2
            dtgNewLx.Text = QBZE
            dtgNewLx.CellForeColor = &H8000&
            htQY = Val(txtHtze.Text) - QBZE
            Exit For
        End If

Next
For oo = 1 To Me.dtgNewLx.Rows
    dtgNewLx.Row = oo
    dtgNewLx.Col = 0
    If dtgNewLx.Text = "" Then
        Exit For
    End If

    dtgNewLx.Col = 4
    If Not (mod1.GxName = "���۹���" And mod1.GXF = True And Me.HTLX = "ѯ��ָ��") Then
        If Trim(dtgNewLx.Text) = "�����" Then
            dtgNewLx.Col = 0
            If Not (InStr(1, dtgNewLx.Text, "Ԥ��") > 0) Then
                dtgNewLx.Col = 1
                Lje = htQY * Val(dtgNewLx.Text) / ZE
                LXG = LXG + Lje
                dtgNewLx.Col = 2: dtgNewLx.Text = Round(Lje, 2): dtgNewLx.CellForeColor = &H8000&
            End If
        End If
    End If
Next
'���һ��ȡ��
dtgNewLx.Col = 4
If Trim(dtgNewLx.Text) = "�����" Then
    If Not (InStr(1, dtgNewLx.Text, "Ԥ��") > 0) Then
        dtgNLN.Row = oo - 1
        LXG = LXG - Lje
        Lje = htQY - LXG
        dtgNewLx.Col = 2: dtgNewLx.Text = Lje: dtgNewLx.CellForeColor = &H8000&
    End If
End If
YGSD = 0
'����Ԥ���ɱ���׼�ۼ�MFϵ��
For oo = 1 To Me.dtgNewLx.Rows
    dtgNewLx.Row = oo
    dtgNewLx.Col = 0
    If dtgNewLx.Text = "" Then
        Exit For
    End If
    
    dtgNewLx.Col = 4
    If Trim(dtgNewLx.Text) = "�����" Then
        dtgNewLx.Col = 0
        If InStr(1, dtgNewLx.Text, "Ԥ��") > 0 Then
            dtgNewLx.Col = 2
            YGSD = Val(dtgNewLx.Text)
            dtgNewLx.Col = 1
            
            dtgNewLx.Text = Round(QBZE / ((Val(txtHtze.Text) - QBZE - Val(txtYJ.Text)) / Val(lblCBZE.Caption)), 2)
            txtQb.Text = Val(dtgNewLx.Text)
            txtQB1.Text = Round(Val(txtQb.Text) / 2.2, 2)
        End If
    End If
Next
End Sub

Private Sub txtKhmc_DblClick()
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


Private Sub txtXmmc_DblClick()
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


Private Sub txtYJ_DblClick()
If lblHTF.ToolTipText = 1 Or lblHTF.ToolTipText = 2 Or lblHTF.ToolTipText = 3 Then
    frmYm.Visible = True
End If
End Sub



Public Sub dtgYjFF()
    dtgYJ.Clear
    dtgYJ.Cols = 4
    dtgYJ.Rows = 12
    dtgYJ.Row = 0
    dtgYJ.Col = 0: dtgYJ.Text = "�տ���": dtgYJ.CellFontBold = True
    dtgYJ.Col = 1: dtgYJ.Text = "����": dtgYJ.CellFontBold = True
    dtgYJ.ColWidth(2) = 0
    dtgYJ.ColWidth(3) = 0
End Sub

Public Sub YjBound(RD, Ld As Integer)
Dim oo As Integer
Call dtgYjFF
For oo = 1 To Ld
    dtgYJ.Row = oo
    dtgYJ.Col = 0: dtgYJ.Text = RD(0, oo - 1)
    dtgYJ.Col = 1: dtgYJ.Text = RD(1, oo - 1)
    dtgYJ.Col = 2: dtgYJ.Text = RD(2, oo - 1)
    dtgYJ.Col = 3: dtgYJ.Text = RD(3, oo - 1)
Next
End Sub

Public Sub HTInput(xZ As Integer)
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

If Lc > 1 And mod1.DName <> "����" And mod1.Mname <> "������" Then Exit Sub

On Error GoTo DER11
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
ElseIf Right(Fname, 4) = ".pdf" Then
    FLX = Right(Fname, 3)
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
    adoFile.Recordset.Update "XZ", xZ
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
    adoFile.Recordset.Update "XZ", xZ '��������
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
MsgBox "�ɹ�����,��ͬ���رգ������������ٴ�!"
If xZ = 0 Then
    cmdDZ.Visible = False
Else
    cmdDz1.Visible = False
End If
If htBrow.Visible = True Then
    frmZu.Enabled = True
    htBrow.Enabled = True
    htBrow.ZOrder 0
ElseIf htBrowG.Visible = True Then
    frmZu.Enabled = True
    htBrowG.Enabled = True
    htBrowG.ZOrder 0
End If
FmxcNew.Visible = False
Exit Sub
DER11:
Close #1
End Sub


Public Sub NewLx()
Dim tt As String
Dim Ra
Dim La As Integer
Dim oo As Integer
dtgNewLx.Clear
dtgNLN.Clear
dtgNewLx.Cols = 9
dtgNewLx.Rows = 20
dtgNewLx.Row = 0
dtgNLN.Cols = 9
dtgNLN.Rows = 20
dtgNLN.Row = 0

dtgNewLx.Col = 0: dtgNewLx.Text = "ҵ������": dtgNewLx.CellFontBold = True
dtgNewLx.Col = 1: dtgNewLx.Text = "��׼��": dtgNewLx.CellFontBold = True
If mod1.GxName = "���۹���" And mod1.GXF = True And Me.HTLX = "ѯ��ָ��" Then
    dtgNewLx.Col = 2: dtgNewLx.Text = "���ⱨ��": dtgNewLx.CellFontBold = True
Else
    dtgNewLx.Col = 2: dtgNewLx.Text = "�ٴ���": dtgNewLx.CellFontBold = True
End If
dtgNewLx.Col = 3: dtgNewLx.Text = "���": dtgNewLx.CellFontBold = True
dtgNewLx.Col = 4: dtgNewLx.Text = "ִ��״̬": dtgNewLx.CellFontBold = True
dtgNewLx.Col = 5: dtgNewLx.Text = "����": dtgNewLx.CellFontBold = True
dtgNewLx.Col = 6: dtgNewLx.Text = "������": dtgNewLx.CellFontBold = True
dtgNewLx.Col = 7: dtgNewLx.Text = "bid": dtgNewLx.CellFontBold = True
dtgNewLx.Col = 8: dtgNewLx.Text = "htrow": dtgNewLx.CellFontBold = True
dtgNewLx.ColWidth(0) = 2280
dtgNewLx.ColWidth(3) = 2580
dtgNewLx.ColWidth(6) = 0
dtgNewLx.ColWidth(7) = 0
dtgNewLx.ColWidth(8) = 0
End Sub

Public Sub LXBound(Rf, Rg)
Dim LF As Integer
Dim Lg As Integer
Dim oo As Integer
Dim tt As String
On Error Resume Next
Call Me.NewLx
If IsNull(Rf(0, 0)) = True Then
    tt = "select zl,jhg,sdje,'BJD'+cast(bid as nvarchar(20)),lc,0,bid,lcren,htrow from xunjiaD where htbh='" & lblHid.Caption & "' and delf=1 order by bid;" & _
            "select dbo.htzui.zl,sum(dbo.htzuidetail.ze) as Ze,0,dbo.htzui.bh,dbo.htzui.lc,1,dbo.htzui.zid,dbo.htzui.lcren,dbo.htzui.htrow" & _
            "  FROM dbo.htZui LEFT OUTER JOIN dbo.htZuiDetail ON dbo.htZui.Zid = dbo.htZuiDetail.Zid where dbo.htzui.hid=" & Val(lblHid.Caption) & " and dbo.htzui.delf=1 " & _
            " group by dbo.htzui.zl,dbo.htzui.bh,dbo.htzui.lc,dbo.htzui.lcren,dbo.htzui.zid,dbo.htzui.htrow order by dbo.htzui.zid"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Rf = mod1.HTP.GetRows
        Set mod1.HTP = mod1.HTP.NextRecordset
        Rg = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
End If


LF = UBound(Rf, 2) + 1
Lg = UBound(Rg, 2) + 1
dtgNewLx.Rows = LF + Lg + 20: dtgNLN.Rows = LF + Lb + 20
'ѯ�۵�
For oo = 1 To LF
    dtgNewLx.Row = oo
    dtgNewLx.Col = 0: dtgNewLx.Text = Rf(0, oo - 1)
    dtgNewLx.Col = 1: dtgNewLx.Text = Rf(1, oo - 1)
    dtgNewLx.Col = 2: dtgNewLx.Text = Rf(2, oo - 1)
    dtgNewLx.Col = 3: dtgNewLx.Text = Rf(3, oo - 1)

    dtgNewLx.Col = 4: dtgNewLx.Text = Rf(4, oo - 1)

        If Rf(4, oo - 1) <> 100 Then
            dtgNewLx.Text = "����" & Rf(7, oo - 1)
            dtgNewLx.Col = 1: dtgNewLx.Text = "" '�������û�н����������ֻ�׼�ۺ��ٴ���
            If Not (mod1.GxName = "���۹���" And mod1.GXF = True And Me.HTLX = "ѯ��ָ��") Then
                dtgNewLx.Col = 2: dtgNewLx.Text = ""
            End If
        Else
            dtgNewLx.Text = "����ˡ�"
        End If

    dtgNewLx.Col = 5: dtgNewLx.Text = Rf(5, oo - 1)
    If dtgNewLx.Text = 0 Then dtgNewLx.Text = "ѯ�۵�"
    dtgNewLx.Col = 6: dtgNewLx.Text = Rf(6, oo - 1)
    dtgNewLx.Col = 7: dtgNewLx.Text = Rf(7, oo - 1)
    dtgNewLx.Col = 8: dtgNewLx.Text = Rf(8, oo - 1)
    
    dtgNLN.Row = oo
    dtgNLN.Col = 0: dtgNLN.Text = Rf(0, oo - 1)
    dtgNLN.Col = 1: dtgNLN.Text = Rf(1, oo - 1)
    dtgNLN.Col = 2: dtgNLN.Text = Rf(2, oo - 1)
    dtgNLN.Col = 3: dtgNLN.Text = Rf(3, oo - 1)

    dtgNLN.Col = 4: dtgNLN.Text = Rf(4, oo - 1)
    If Rf(4, oo - 1) <> 100 Then
        dtgNLN.Text = "����" & Rf(7, oo - 1)
    Else
        dtgNLN.Text = "�����"
    End If
    dtgNLN.Col = 5: dtgNLN.Text = Rf(5, oo - 1)
    If dtgNLN.Text = 0 Then dtgNLN.Text = "ѯ�۵�"
    dtgNLN.Col = 6: dtgNLN.Text = Rf(6, oo - 1)
    dtgNLN.Col = 7: dtgNLN.Text = Rf(7, oo - 1)
    dtgNLN.Col = 8: dtgNLN.Text = Rf(8, oo - 1)
Next

''''''''''�ɱ������
'''''''''For oo = 1 To Lg
'''''''''    dtgNewLx.Row = oo + LF
'''''''''    dtgNewLx.Col = 0: dtgNewLx.Text = Rg(0, oo - 1)
'''''''''    dtgNewLx.Col = 1: dtgNewLx.Text = Rg(1, oo - 1)
'''''''''    dtgNewLx.Col = 2: dtgNewLx.Text = Rg(2, oo - 1)
'''''''''    dtgNewLx.Col = 3: dtgNewLx.Text = Rg(3, oo - 1)
'''''''''
'''''''''    dtgNewLx.Col = 4: dtgNewLx.Text = Rg(4, oo - 1)
'''''''''    If Rg(4, oo - 1) <> 100 Then
'''''''''        dtgNewLx.Text = "����" & Rg(7, oo - 1)
'''''''''        dtgNewLx.Col = 1: dtgNewLx.Text = "" '�������û�н����������ֻ�׼�ۺ��ٴ���
'''''''''        dtgNewLx.Col = 2: dtgNewLx.Text = ""
'''''''''    Else
'''''''''        dtgNewLx.Text = "����ˡ�"
'''''''''    End If
'''''''''    dtgNewLx.Col = 5: dtgNewLx.Text = Rg(5, oo - 1)
'''''''''    dtgNewLx.Text = "�ɱ������"
'''''''''    If dtgNewLx.Text = 0 Then dtgNewLx.Text = "�ɱ������"
'''''''''    dtgNewLx.Col = 6: dtgNewLx.Text = Rg(6, oo - 1)
'''''''''    dtgNewLx.Col = 7: dtgNewLx.Text = Rg(7, oo - 1)
'''''''''    dtgNewLx.Col = 8: dtgNewLx.Text = Rg(8, oo - 1)
'''''''''
'''''''''    dtgNLN.Row = oo + LF
'''''''''    dtgNLN.Col = 0: dtgNLN.Text = Rg(0, oo - 1)
'''''''''    dtgNLN.Col = 1: dtgNLN.Text = Rg(1, oo - 1)
'''''''''    dtgNLN.Col = 2: dtgNLN.Text = Rg(2, oo - 1)
'''''''''    dtgNLN.Col = 3: dtgNLN.Text = Rg(3, oo - 1)
'''''''''
'''''''''    dtgNLN.Col = 4: dtgNLN.Text = Rg(4, oo - 1)
'''''''''    If Rg(4, oo - 1) <> 100 Then
'''''''''        dtgNLN.Text = "����" & Rg(7, oo - 1)
'''''''''        dtgNLN.Col = 1: dtgNLN.Text = "" '�������û�н����������ֻ�׼�ۺ��ٴ���
'''''''''        dtgNLN.Col = 2: dtgNLN.Text = ""
'''''''''    Else
'''''''''        dtgNLN.Text = "����ˡ�"
'''''''''    End If
'''''''''    dtgNLN.Col = 5: dtgNLN.Text = Rg(5, oo - 1)
'''''''''    dtgNLN.Text = "�ɱ������"
'''''''''    If dtgNLN.Text = 0 Then dtgNLN.Text = "�ɱ������"
'''''''''    dtgNLN.Col = 6: dtgNLN.Text = Rg(6, oo - 1)
'''''''''    dtgNLN.Col = 7: dtgNLN.Text = Rg(7, oo - 1)
'''''''''    dtgNLN.Col = 8: dtgNLN.Text = Rg(8, oo - 1)
'''''''''Next
End Sub
Public Sub LXBound1(Rf, Rg)
Dim LF As Integer
Dim Lg As Integer
Dim oo As Integer
Dim tt As String
On Error Resume Next
Call Me.NewLx
If IsNull(Rf(0, 0)) = True Then
    tt = "select dbo.htzui.zl,sum(dbo.htzuidetail.ze) as Ze,0,dbo.htzui.bh,dbo.htzui.lc,1,dbo.htzui.zid,dbo.htzui.lcren,dbo.htzui.htrow" & _
            "  FROM dbo.htZui LEFT OUTER JOIN dbo.htZuiDetail ON dbo.htZui.Zid = dbo.htZuiDetail.Zid where dbo.htzui.hid=" & Val(lblHid.Caption) & " and dbo.htzui.delf=1 " & _
            " and dbo.htzui.fl='׷��' group by dbo.htzui.zl,dbo.htzui.bh,dbo.htzui.lc,dbo.htzui.lcren,dbo.htzui.zid,dbo.htzui.htrow order by dbo.htzui.zid"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Rg = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
End If

Lg = UBound(Rg, 2) + 1
dtgNewLx.Rows = LF + Lg + 20: dtgNLN.Rows = dtgNewLx.Rows
''''''ѯ�۵�
'''''For oo = 1 To LF
'''''    dtgNewLx.Row = oo
'''''    dtgNewLx.Col = 0: dtgNewLx.Text = Rf(0, oo - 1)
'''''    dtgNewLx.Col = 1: dtgNewLx.Text = Rf(1, oo - 1)
'''''    dtgNewLx.Col = 2: dtgNewLx.Text = Rf(2, oo - 1)
'''''    dtgNewLx.Col = 3: dtgNewLx.Text = Rf(3, oo - 1)
'''''
'''''    dtgNewLx.Col = 4: dtgNewLx.Text = Rf(4, oo - 1)
'''''    If Rf(4, oo - 1) <> 100 Then
'''''        dtgNewLx.Text = "����" & Rf(7, oo - 1)
'''''        dtgNewLx.Col = 1: dtgNewLx.Text = "" '�������û�н����������ֻ�׼�ۺ��ٴ���
'''''        dtgNewLx.Col = 2: dtgNewLx.Text = ""
'''''    Else
'''''        dtgNewLx.Text = "����ˡ�"
'''''    End If
'''''    dtgNewLx.Col = 5: dtgNewLx.Text = Rf(5, oo - 1)
'''''    If dtgNewLx.Text = 0 Then dtgNewLx.Text = "ѯ�۵�"
'''''    dtgNewLx.Col = 6: dtgNewLx.Text = Rf(6, oo - 1)
'''''    dtgNewLx.Col = 7: dtgNewLx.Text = Rf(7, oo - 1)
'''''    dtgNewLx.Col = 8: dtgNewLx.Text = Rf(8, oo - 1)
'''''
'''''    dtgNLN.Row = oo
'''''    dtgNLN.Col = 0: dtgNLN.Text = Rf(0, oo - 1)
'''''    dtgNLN.Col = 1: dtgNLN.Text = Rf(1, oo - 1)
'''''    dtgNLN.Col = 2: dtgNLN.Text = Rf(2, oo - 1)
'''''    dtgNLN.Col = 3: dtgNLN.Text = Rf(3, oo - 1)
'''''
'''''    dtgNLN.Col = 4: dtgNLN.Text = Rf(4, oo - 1)
'''''    If Rf(4, oo - 1) <> 100 Then
'''''        dtgNLN.Text = "����" & Rf(7, oo - 1)
'''''    Else
'''''        dtgNLN.Text = "�����"
'''''    End If
'''''    dtgNLN.Col = 5: dtgNLN.Text = Rf(5, oo - 1)
'''''    If dtgNLN.Text = 0 Then dtgNLN.Text = "ѯ�۵�"
'''''    dtgNLN.Col = 6: dtgNLN.Text = Rf(6, oo - 1)
'''''    dtgNLN.Col = 7: dtgNLN.Text = Rf(7, oo - 1)
'''''    dtgNLN.Col = 8: dtgNLN.Text = Rf(8, oo - 1)
'''''Next
'Call NewCale
'�ɱ������
For oo = 1 To Lg
    dtgNewLx.Row = oo
    dtgNewLx.Col = 0: dtgNewLx.Text = Rg(0, oo - 1)
    dtgNewLx.Col = 1: dtgNewLx.Text = Rg(1, oo - 1)
    dtgNewLx.Col = 2: dtgNewLx.Text = Rg(2, oo - 1)
    dtgNewLx.Col = 3: dtgNewLx.Text = Rg(3, oo - 1)

    dtgNewLx.Col = 4: dtgNewLx.Text = Rg(4, oo - 1)
    If Rg(4, oo - 1) <> 100 Then
        dtgNewLx.Text = "����" & Rg(7, oo - 1)
        dtgNewLx.Col = 1: dtgNewLx.Text = "" '�������û�н����������ֻ�׼�ۺ��ٴ���
        dtgNewLx.Col = 2: dtgNewLx.Text = ""
    Else
        dtgNewLx.Text = "����ˡ�"
    End If
    dtgNewLx.Col = 5: dtgNewLx.Text = Rg(5, oo - 1)
    dtgNewLx.Text = "�ɱ������"
    If dtgNewLx.Text = 0 Then dtgNewLx.Text = "�ɱ������"
    dtgNewLx.Col = 6: dtgNewLx.Text = Rg(6, oo - 1)
    dtgNewLx.Col = 7: dtgNewLx.Text = Rg(7, oo - 1)
    dtgNewLx.Col = 8: dtgNewLx.Text = Rg(8, oo - 1)

    dtgNLN.Row = oo
    dtgNLN.Col = 0: dtgNLN.Text = Rg(0, oo - 1)
    dtgNLN.Col = 1: dtgNLN.Text = Rg(1, oo - 1)
    dtgNLN.Col = 2: dtgNLN.Text = Rg(2, oo - 1)
    dtgNLN.Col = 3: dtgNLN.Text = Rg(3, oo - 1)

    dtgNLN.Col = 4: dtgNLN.Text = Rg(4, oo - 1)
    If Rg(4, oo - 1) <> 100 Then
        dtgNLN.Text = "����" & Rg(7, oo - 1)
        dtgNLN.Col = 1: dtgNLN.Text = "" '�������û�н����������ֻ�׼�ۺ��ٴ���
        dtgNLN.Col = 2: dtgNLN.Text = ""
    Else
        dtgNLN.Text = "����ˡ�"
    End If
    dtgNLN.Col = 5: dtgNLN.Text = Rg(5, oo - 1)
    dtgNLN.Text = "�ɱ������"
    If dtgNLN.Text = 0 Then dtgNLN.Text = "�ɱ������"
    dtgNLN.Col = 6: dtgNLN.Text = Rg(6, oo - 1)
    dtgNLN.Col = 7: dtgNLN.Text = Rg(7, oo - 1)
    dtgNLN.Col = 8: dtgNLN.Text = Rg(8, oo - 1)
Next
dtgNewLx.Row = 0
End Sub


Public Function JCYG() As Boolean '����Ƿ��г���һ���Ԥ���ɱ�
Dim oo As Integer
Dim ii As Integer
On Error Resume Next
ii = 0
JCYG = False
YGCB = 0
For oo = 1 To 100
    dtgNewLx.Col = 0
    dtgNewLx.Row = oo
    If dtgNewLx.Text = "" Then Exit For
    If InStr(1, dtgNewLx.Text, "Ԥ��") > 0 Then
        dtgNewLx.Col = 1
        YGCB = Val(dtgNewLx.Text)
        ii = ii + 1
    End If
Next
If ii > 1 Then
    JCYG = True
End If
    
End Function

Public Sub LXBound2(Rf, Rg)
Dim LF As Integer
Dim Lg As Integer
Dim oo As Integer
Dim tt As String
On Error Resume Next
Call Me.NewLx
If IsNull(Rf(0, 0)) = True Then
    tt = "select dbo.htzui.zl,sum(dbo.htzuidetail.ze) as Ze,0,dbo.htzui.bh,dbo.htzui.lc,1,dbo.htzui.zid,dbo.htzui.lcren,dbo.htzui.htrow" & _
            "  FROM dbo.htZui LEFT OUTER JOIN dbo.htZuiDetail ON dbo.htZui.Zid = dbo.htZuiDetail.Zid where dbo.htzui.hid=" & Val(lblHid.Caption) & " and dbo.htzui.delf=1 " & _
            " and dbo.htzui.fl='����' group by dbo.htzui.zl,dbo.htzui.bh,dbo.htzui.lc,dbo.htzui.lcren,dbo.htzui.zid,dbo.htzui.htrow order by dbo.htzui.zid"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Rg = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
End If

Lg = UBound(Rg, 2) + 1
dtgNewLx.Rows = LF + Lg + 20: dtgNLN.Rows = dtgNewLx.Rows


'�ɱ������
For oo = 1 To Lg
    dtgNewLx.Row = oo
    dtgNewLx.Col = 0: dtgNewLx.Text = Rg(0, oo - 1)
    dtgNewLx.Col = 1: dtgNewLx.Text = Rg(1, oo - 1)
    dtgNewLx.Col = 2: dtgNewLx.Text = Rg(2, oo - 1)
    dtgNewLx.Col = 3: dtgNewLx.Text = Rg(3, oo - 1)

    dtgNewLx.Col = 4: dtgNewLx.Text = Rg(4, oo - 1)
    If Rg(4, oo - 1) <> 100 Then
        dtgNewLx.Text = "����" & Rg(7, oo - 1)
        dtgNewLx.Col = 1: dtgNewLx.Text = "" '�������û�н����������ֻ�׼�ۺ��ٴ���
        dtgNewLx.Col = 2: dtgNewLx.Text = ""
    Else
        dtgNewLx.Text = "����ˡ�"
    End If
    dtgNewLx.Col = 5: dtgNewLx.Text = Rg(5, oo - 1)
    dtgNewLx.Text = "�ɱ������"
    If dtgNewLx.Text = 0 Then dtgNewLx.Text = "�ɱ������"
    dtgNewLx.Col = 6: dtgNewLx.Text = Rg(6, oo - 1)
    dtgNewLx.Col = 7: dtgNewLx.Text = Rg(7, oo - 1)
    dtgNewLx.Col = 8: dtgNewLx.Text = Rg(8, oo - 1)

    dtgNLN.Row = oo
    dtgNLN.Col = 0: dtgNLN.Text = Rg(0, oo - 1)
    dtgNLN.Col = 1: dtgNLN.Text = Rg(1, oo - 1)
    dtgNLN.Col = 2: dtgNLN.Text = Rg(2, oo - 1)
    dtgNLN.Col = 3: dtgNLN.Text = Rg(3, oo - 1)

    dtgNLN.Col = 4: dtgNLN.Text = Rg(4, oo - 1)
    If Rg(4, oo - 1) <> 100 Then
        dtgNLN.Text = "����" & Rg(7, oo - 1)
        dtgNLN.Col = 1: dtgNLN.Text = "" '�������û�н����������ֻ�׼�ۺ��ٴ���
        dtgNLN.Col = 2: dtgNLN.Text = ""
    Else
        dtgNLN.Text = "����ˡ�"
    End If
    dtgNLN.Col = 5: dtgNLN.Text = Rg(5, oo - 1)
    dtgNLN.Text = "�ɱ������"
    If dtgNLN.Text = 0 Then dtgNLN.Text = "�ɱ������"
    dtgNLN.Col = 6: dtgNLN.Text = Rg(6, oo - 1)
    dtgNLN.Col = 7: dtgNLN.Text = Rg(7, oo - 1)
    dtgNLN.Col = 8: dtgNLN.Text = Rg(8, oo - 1)
Next
dtgNewLx.Row = 0
End Sub

Public Sub Xian()
Dim oo As Long
On Error Resume Next
            FmxcNew.txtHtbh.Top = 100
            FmxcNew.lblHtbh.Top = 100
            FmxcNew.lblHtrq.Visible = False
            FmxcNew.txtHtrq.Visible = False
            FmxcNew.dtgFk.Top = 500
            FmxcNew.dtgFk.ColWidth(0) = 3000
            FmxcNew.dtgFk.ColWidth(2) = 0
            FmxcNew.dtgFk.Height = 7500
            For oo = 0 To 50
                FmxcNew.dtgNewLx.Row = oo
                FmxcNew.dtgNewLx.Col = 2
                FmxcNew.dtgNewLx.Text = ""
                FmxcNew.dtgNewLx.ColWidth(2) = 0
            Next
End Sub
