VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form fmxcZJ 
   BackColor       =   &H00C0FFC0&
   Caption         =   "׷�ӳɱ���"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "���"
      Height          =   765
      Left            =   11490
      Picture         =   "fmxcZJ.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   8220
      Width           =   675
   End
   Begin VB.Frame frmCg 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�༭"
      Height          =   1875
      Left            =   0
      TabIndex        =   78
      Top             =   5310
      Width           =   5175
      Begin VB.TextBox txtSL 
         Height          =   270
         Left            =   2910
         TabIndex        =   90
         Top             =   390
         Width           =   1125
      End
      Begin VB.CommandButton cmdDao 
         BackColor       =   &H00FFFF00&
         Caption         =   "��Ʒ���"
         Height          =   345
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   870
         Width           =   915
      End
      Begin VB.CommandButton cmdNGx 
         BackColor       =   &H00FF8080&
         Caption         =   "����"
         Height          =   345
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   870
         Width           =   855
      End
      Begin VB.CommandButton cmdNDel 
         BackColor       =   &H008080FF&
         Caption         =   "����"
         Height          =   345
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   870
         Width           =   855
      End
      Begin VB.Frame frmJ 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   285
         Left            =   -150
         TabIndex        =   80
         Top             =   360
         Width           =   2235
         Begin VB.TextBox txtJdj 
            Height          =   270
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   30
            Width           =   1155
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "��׼����"
            Height          =   255
            Left            =   300
            TabIndex        =   82
            Top             =   60
            Width           =   855
         End
      End
      Begin VB.Frame frmZ 
         Height          =   405
         Left            =   -8310
         TabIndex        =   84
         Top             =   690
         Width           =   8295
      End
      Begin VB.TextBox txtDj 
         Height          =   270
         Left            =   930
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   390
         Width           =   1155
      End
      Begin VB.CommandButton cmdGy 
         BackColor       =   &H00C0E0FF&
         Caption         =   "��Ӧ��"
         Height          =   315
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   900
         Width           =   885
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   225
         Index           =   1
         Left            =   2310
         TabIndex        =   91
         Top             =   420
         Width           =   375
      End
      Begin VB.Label lblDj 
         BackStyle       =   0  'Transparent
         Caption         =   "�ɱ�����"
         Height          =   195
         Left            =   120
         TabIndex        =   86
         Top             =   450
         Width           =   765
      End
      Begin VB.Label lblDid 
         Caption         =   "lblDid"
         Height          =   255
         Left            =   3150
         TabIndex        =   85
         Top             =   930
         Visible         =   0   'False
         Width           =   825
      End
   End
   Begin VB.Frame frmGY 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��Ӧ�̼۸�"
      Height          =   1995
      Left            =   5160
      TabIndex        =   63
      Top             =   5310
      Visible         =   0   'False
      Width           =   7245
      Begin VB.TextBox txtGy1 
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "Text1"
         Top             =   390
         Width           =   3195
      End
      Begin VB.TextBox txtGy2 
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   825
         Width           =   3195
      End
      Begin VB.TextBox txtGY3 
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text1"
         Top             =   1260
         Width           =   3195
      End
      Begin VB.TextBox txtGdj1 
         Height          =   270
         Left            =   5280
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   390
         Width           =   765
      End
      Begin VB.TextBox txtGdj2 
         Height          =   285
         Left            =   5280
         TabIndex        =   69
         Text            =   "Text2"
         Top             =   787
         Width           =   765
      End
      Begin VB.TextBox txtGdj3 
         Height          =   285
         Left            =   5280
         TabIndex        =   68
         Text            =   "Text3"
         Top             =   1230
         Width           =   765
      End
      Begin VB.OptionButton optGy1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "��Ӧ��1"
         Height          =   285
         Left            =   180
         TabIndex        =   67
         Top             =   390
         Width           =   975
      End
      Begin VB.OptionButton optGy2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "��Ӧ��2"
         Height          =   285
         Left            =   180
         TabIndex        =   66
         Top             =   810
         Width           =   975
      End
      Begin VB.OptionButton optGy3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "��Ӧ��3"
         Height          =   285
         Left            =   180
         TabIndex        =   65
         Top             =   1230
         Width           =   975
      End
      Begin VB.TextBox txtGy 
         Height          =   315
         Left            =   6180
         TabIndex        =   64
         Top             =   1560
         Width           =   3735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgGy 
         Height          =   1335
         Left            =   6180
         TabIndex        =   74
         Top             =   120
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   2355
         _Version        =   393216
         BackColor       =   12648384
         Rows            =   50
         FixedCols       =   0
         BackColorFixed  =   12648384
         BackColorBkg    =   16777152
         WordWrap        =   -1  'True
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         PictureType     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "�۸�1"
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   77
         Top             =   420
         Width           =   525
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "�۸�2"
         Height          =   255
         Left            =   4680
         TabIndex        =   76
         Top             =   825
         Width           =   525
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "�۸�3"
         Height          =   255
         Left            =   4680
         TabIndex        =   75
         Top             =   1260
         Width           =   525
      End
   End
   Begin VB.Frame frmGui 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   465
      Left            =   12330
      TabIndex        =   53
      Top             =   600
      Width           =   2865
      Begin VB.ComboBox comGui 
         ForeColor       =   &H000000FF&
         Height          =   300
         ItemData        =   "fmxcZJ.frx":0442
         Left            =   870
         List            =   "fmxcZJ.frx":0452
         TabIndex        =   54
         Top             =   0
         Width           =   1725
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "�ɱ�����"
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   0
         TabIndex        =   55
         Top             =   30
         Width           =   735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   435
      Left            =   3420
      TabIndex        =   50
      Top             =   5520
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   767
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "������"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   -30
      TabIndex        =   36
      Top             =   7320
      Visible         =   0   'False
      Width           =   6315
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "����"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "�ܾ�"
         Height          =   195
         Left            =   5220
         TabIndex        =   39
         Top             =   870
         Width           =   675
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ͬ��"
         Height          =   225
         Left            =   5220
         TabIndex        =   38
         Top             =   510
         Width           =   705
      End
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1305
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   300
         Width           =   4965
      End
   End
   Begin VB.Frame frmBan 
      BackColor       =   &H00FFFFC0&
      Height          =   2955
      Left            =   8760
      TabIndex        =   22
      Top             =   2340
      Width           =   6075
      Begin VB.Frame frmDj 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   405
         Left            =   2310
         TabIndex        =   43
         Top             =   2070
         Width           =   3735
         Begin VB.TextBox txt6 
            Height          =   270
            Left            =   630
            TabIndex        =   45
            Text            =   "Text6"
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox txt7 
            Height          =   285
            Left            =   2640
            TabIndex        =   44
            Text            =   "Text7"
            Top             =   0
            Width           =   945
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   255
            Left            =   0
            TabIndex        =   47
            Top             =   30
            Width           =   465
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "�ܼ�"
            Height          =   255
            Left            =   2160
            TabIndex        =   46
            Top             =   30
            Width           =   405
         End
      End
      Begin VB.CommandButton cmdBG 
         Caption         =   "�ر�"
         Height          =   315
         Left            =   4950
         TabIndex        =   42
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdD 
         Caption         =   "ɾ��"
         Height          =   315
         Left            =   3380
         TabIndex        =   35
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdGx 
         Caption         =   "����"
         Height          =   315
         Left            =   1810
         TabIndex        =   34
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "���"
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txt5 
         Height          =   270
         Left            =   870
         TabIndex        =   32
         Text            =   "Text5"
         Top             =   2070
         Width           =   795
      End
      Begin VB.TextBox txt4 
         Height          =   315
         Left            =   2310
         TabIndex        =   30
         Text            =   "Text4"
         Top             =   1605
         Width           =   3615
      End
      Begin VB.TextBox txt3 
         Height          =   315
         Left            =   2310
         TabIndex        =   28
         Text            =   "Text3"
         Top             =   1125
         Width           =   3615
      End
      Begin VB.TextBox txt2 
         Height          =   315
         Left            =   2310
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   660
         Width           =   3585
      End
      Begin VB.TextBox txt1 
         Height          =   315
         Left            =   2310
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   180
         Width           =   3585
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   255
         Left            =   180
         TabIndex        =   31
         Top             =   2100
         Width           =   615
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ӧ/�ְ���"
         Height          =   255
         Left            =   150
         TabIndex        =   29
         Top             =   1635
         Width           =   1755
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "����ͺ�"
         Height          =   255
         Left            =   150
         TabIndex        =   27
         Top             =   1155
         Width           =   1605
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "����Ʒ��(�ͺ�)"
         Height          =   255
         Left            =   150
         TabIndex        =   25
         Top             =   660
         Width           =   1725
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "��/������/��������"
         Height          =   255
         Left            =   150
         TabIndex        =   23
         Top             =   180
         Width           =   1785
      End
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   420
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   420
      Top             =   0
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "�ύ"
      Height          =   765
      Left            =   12870
      Picture         =   "fmxcZJ.frx":048E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8220
      Width           =   675
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   765
      Left            =   14220
      Picture         =   "fmxcZJ.frx":0AF8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8220
      Width           =   585
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "�޸�"
      Height          =   765
      Left            =   12210
      Picture         =   "fmxcZJ.frx":0BFA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8220
      Width           =   645
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ��"
      Enabled         =   0   'False
      Height          =   765
      Left            =   13560
      Picture         =   "fmxcZJ.frx":103C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8220
      Width           =   645
   End
   Begin VB.TextBox txtBz 
      Height          =   1395
      Left            =   10500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "fmxcZJ.frx":11C6
      Top             =   5730
      Width           =   4545
   End
   Begin VB.OptionButton optF 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��ȫ����ͬ"
      Height          =   300
      Left            =   13530
      TabIndex        =   3
      Top             =   180
      Width           =   1545
   End
   Begin VB.OptionButton optQ 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ȫ����ͬ"
      Height          =   300
      Left            =   12360
      TabIndex        =   2
      Top             =   180
      Width           =   1035
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMx 
      Height          =   3945
      Left            =   180
      TabIndex        =   13
      Top             =   1350
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   6959
      _Version        =   393216
      BackColor       =   12648384
      BackColorFixed  =   12648384
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   3375
      Left            =   0
      TabIndex        =   14
      Top             =   5700
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   5953
      _Version        =   393216
      BackColor       =   12648447
      ForeColor       =   8404992
      Rows            =   15
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   16761024
      ForeColorFixed  =   0
      BackColorBkg    =   12648447
      GridColorFixed  =   8404992
      GridColorUnpopulated=   8404992
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label lblMF 
      BackStyle       =   0  'Transparent
      Caption         =   "Label25"
      Height          =   255
      Left            =   9720
      TabIndex        =   94
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MFϵ��"
      Height          =   255
      Left            =   8880
      TabIndex        =   93
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label lblCB2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label22"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3900
      TabIndex        =   62
      Top             =   990
      Width           =   1305
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "�ۼ��ܶ�"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3060
      TabIndex        =   61
      Top             =   990
      Width           =   915
   End
   Begin VB.Label lblCb1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1860
      TabIndex        =   60
      Top             =   990
      Width           =   1095
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "��ͬԤ���ɱ��ܶ�"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   270
      TabIndex        =   59
      Top             =   990
      Width           =   1485
   End
   Begin VB.Label lblZtime 
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   9600
      TabIndex        =   58
      Top             =   210
      Width           =   2535
   End
   Begin VB.Label lblFBF 
      BackStyle       =   0  'Transparent
      Caption         =   "Label19"
      Height          =   255
      Left            =   10740
      TabIndex        =   57
      Top             =   630
      Width           =   1455
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "׷�ӵ�����"
      Height          =   255
      Left            =   9630
      TabIndex        =   56
      Top             =   630
      Width           =   1005
   End
   Begin VB.Label lblYwy 
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      Height          =   225
      Left            =   8370
      TabIndex        =   52
      Top             =   180
      Width           =   825
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "�Ƶ���"
      Height          =   225
      Left            =   7710
      TabIndex        =   51
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblJe 
      BackStyle       =   0  'Transparent
      Caption         =   "Label17"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   7530
      TabIndex        =   49
      Top             =   990
      Width           =   825
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�����ɱ��ܶ�"
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   6150
      TabIndex        =   48
      Top             =   1020
      Width           =   1155
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
      Height          =   345
      Left            =   10410
      TabIndex        =   41
      Top             =   7500
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "����״̬"
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   270
      TabIndex        =   21
      Top             =   5430
      Width           =   1005
   End
   Begin VB.Label lblBh 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   6510
      TabIndex        =   20
      Top             =   180
      Width           =   615
   End
   Begin VB.Label lblZe 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      Height          =   300
      Left            =   8370
      TabIndex        =   19
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label lblXz 
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      Height          =   300
      Left            =   4470
      TabIndex        =   18
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label lblZbh 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      Height          =   300
      Left            =   1320
      TabIndex        =   17
      Top             =   600
      Width           =   2565
   End
   Begin VB.Label lblGLBH 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      Height          =   300
      Left            =   4680
      TabIndex        =   16
      Top             =   180
      Width           =   1815
   End
   Begin VB.Label lblKhmc 
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      Height          =   300
      Left            =   1320
      TabIndex        =   15
      Top             =   180
      Width           =   2565
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "����ԭ��"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   10560
      TabIndex        =   7
      Top             =   5460
      Width           =   1065
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "��ͬ�ܶ�"
      Height          =   300
      Left            =   7530
      TabIndex        =   6
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "��ͬ����"
      Height          =   300
      Left            =   3540
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "��ִͬ�к�"
      Height          =   300
      Left            =   210
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblZid 
      BackStyle       =   0  'Transparent
      Caption         =   "������ͬ���"
      Height          =   210
      Left            =   3540
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ�����"
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "fmxcZJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer '�ɱ�׷�ӵ�(1����2ɾ��3��ϸ�༭5ǩ��19ִ֪ͨ��)

Dim LCRen As String
Dim LCUid As String
Public Lc As Integer
Dim Fwid As Long
Dim xZ As Boolean '�������(0׷�ص�1���˵��)
Public Ywy As String
Public Uid As String
Public htRow As Integer
Dim NewF As Integer '����Ӧ��ͬ�İ汾
Dim Hid As Long '
Dim NewFZJ As Integer
Dim GyId As Integer

Private Sub cmdAdd_Click()
If Val(lblZid.ToolTipText) = 0 Then
    MsgBox "���ȱ���,�������ϸ!"
    Exit Sub
End If
If Val(txt5.Text) = 0 Then
    MsgBox "��ȷ������!"
    Exit Sub
End If


timZm = 3 '����
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�ɱ�׷�ӵ�"
    mod1.cmd.Parameters("@NBLX") = "��ϸ�༭"
    mod1.cmd.Parameters("@bh") = lblZid.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txt1.Text
    mod1.cmd.Parameters("@mt2") = txt2.Text
    mod1.cmd.Parameters("@mt3") = txt3.Text
    mod1.cmd.Parameters("@mt4") = txt4.Text
    mod1.cmd.Parameters("@mt20") = "���"

    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt5.Text) '����
    mod1.cmd.Parameters("@mm2") = Val(txt6.Text) '����
    mod1.cmd.Parameters("@mm3") = Val(txt7.Text) '����


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

        
    End If

    
Set mod1.cmd = Nothing
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub

Private Sub cmdBG_Click()
frmBan.Visible = False
End Sub

Private Sub cmdD_Click()
Dim Did As Long
Dim ii As Integer
Did = Val(txt7.ToolTipText)
If Did = 0 Then Exit Sub

ii = MsgBox("�Ƿ�ɾ���˳ɱ�׷�ӵ�?", vbYesNo + vbQuestion, "��ȷ��")
If ii = vbNo Then Exit Sub


timZm = 3 '
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�ɱ�׷�ӵ�"
    mod1.cmd.Parameters("@NBLX") = "��ϸ�༭"
    mod1.cmd.Parameters("@bh") = lblZid.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txt1.Text
    mod1.cmd.Parameters("@mt2") = txt2.Text
    mod1.cmd.Parameters("@mt3") = txt3.Text
    mod1.cmd.Parameters("@mt4") = txt4.Text
    mod1.cmd.Parameters("@mt20") = "ɾ��"

    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt5.Text) '����
    mod1.cmd.Parameters("@mm2") = Val(txt6.Text) '����
    mod1.cmd.Parameters("@mm3") = Val(txt7.Text) '����
    mod1.cmd.Parameters("@mm20") = Did

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

        
    End If

    
Set mod1.cmd = Nothing
End Sub

Private Sub cmdDao_Click()
Dim tt As String
    tt = "select top 50 bh,partname,'ԭ�����:'+oname+' '+gg+' '+xn+' '+ff+' ���û���:'+jz from nlpmxc order by bh desc"
    Call FmxcXjHp.Bound(tt)
    FmxcXjHp.Show
    FmxcXjHp.ZOrder 0
    If InStr(1, lblFBF.Caption, "�ְ�") > 0 Then
        FmxcXjHp.cmdDao.Caption = "�ְ�����"
    Else
        FmxcXjHp.cmdDao.Caption = "����"
    End If
End Sub

Private Sub cmdDel_Click()
Dim ii As Integer
If lblBh.ToolTipText = "" Then Exit Sub
If Me.Lc > 1 And mod1.DName <> "���ĳ�" And mod1.DName <> "������" Then
    Exit Sub
End If
ii = MsgBox("�Ƿ�ɾ���˳ɱ�׷�ӵ�?", vbYesNo + vbQuestion, "��ȷ��")
If ii = vbNo Then Exit Sub

timZm = 2 'ɾ��
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�ɱ�׷�ӵ�"
    mod1.cmd.Parameters("@NBLX") = "ɾ��"
    mod1.cmd.Parameters("@bh") = Trim(lblBh.ToolTipText)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""

    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(lblGLBH.ToolTipText)  'Hid
    If optQ.Value = True Then
        mod1.cmd.Parameters("@mb1") = 1 'ȫ����
    Else
        mod1.cmd.Parameters("@mb1") = 0 'ȫ����
    End If
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

Private Sub cmdDing_Click()
Dim tt As String
On Error Resume Next
If Lc = 0 Then
    Exit Sub
End If
'����Ƿ񳬳�Ԥ���ɱ�
If mod1.ZT = "HMData" And (NewF = 6 Or NewF = 7 Or NewF = 8) And OptT1.Value = True Then
    If Val(lblCB2.Caption) > Val(lblCb1.Caption) Then
        MsgBox ("����Ԥ���ɱ���")
        Exit Sub
    End If

End If

If comGui.Text = "" Then
    MsgBox "��ȷ�Ϸ��ù���!"
    Exit Sub
End If
If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("����һ��Ҫ���߾ܾ��ҵ�����!  :) ")
    Exit Sub
End If
frmQm.Visible = False
        timZm = 5 'ǩ��
        Set mod1.cmd = New ADODB.command
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "�ɱ�׷�ӵ�"
        mod1.cmd.Parameters("@NBLX") = "ǩ��"
        mod1.cmd.Parameters("@bh") = lblZid.ToolTipText
        If mod1.cmd.Parameters("@bh").Value = 0 Then
            MsgBox ("����!,�����´�����һ��!")
            Me.Visible = False
            Exit Sub
        End If
        
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = comGui.Text
        mod1.cmd.Parameters("@mt2") = mod1.Qy
        mod1.cmd.Parameters("@mt3") = lblKhmc.Caption
        mod1.cmd.Parameters("@mt4") = ""
        mod1.cmd.Parameters("@mt5") = lblYwy.Caption
        mod1.cmd.Parameters("@mt6") = lblYwy.ToolTipText
        mod1.cmd.Parameters("@mt7") = lblFBF.Caption '�ְ��������
        mod1.cmd.Parameters("@mlt1") = txtQM.Text '������

        mod1.cmd.Parameters("@mm1").Value = Me.Lc
        mod1.cmd.Parameters("@mm2").Value = Fwid
        mod1.cmd.Parameters("@mm3") = Hid

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
End Sub

Private Sub cmdGx_Click()
Dim Did As Long
Did = Val(txt7.ToolTipText)
If Did = 0 Then Exit Sub

If Val(lblZid.ToolTipText) = 0 Then
    MsgBox "���ȱ���,�������ϸ!"
    Exit Sub
End If
If Val(txt5.Text) = 0 Then
    MsgBox "��ȷ������!"
    Exit Sub
End If


timZm = 3 '����
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�ɱ�׷�ӵ�"
    mod1.cmd.Parameters("@NBLX") = "��ϸ�༭"
    mod1.cmd.Parameters("@bh") = lblZid.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txt1.Text
    mod1.cmd.Parameters("@mt2") = txt2.Text
    mod1.cmd.Parameters("@mt3") = txt3.Text
    mod1.cmd.Parameters("@mt4") = txt4.Text
    mod1.cmd.Parameters("@mt20") = "����"

    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt5.Text) '����
    mod1.cmd.Parameters("@mm2") = Val(txt6.Text) '����
    mod1.cmd.Parameters("@mm3") = Val(txt7.Text) '����
    mod1.cmd.Parameters("@mm20") = Did

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

        
    End If

    
Set mod1.cmd = Nothing
End Sub


Private Sub cmdGy_Click()
frmGy.Visible = True
End Sub

Private Sub cmdMod_Click()
If cmdSave.Enabled = True Then Exit Sub
If mod1.ZT = "HMData" Or mod1.Mname = "������" Then
    If Me.Lc = 1 And mod1.DName = LCRen Then
        cmdDel.Enabled = True
        frmGui.Enabled = True
        optF.Enabled = True
        optQ.Enabled = True
        cmdGy.Visible = False
        frmJ.Visible = True
        cmdSave.Enabled = True
    ElseIf Me.Lc = 2 And mod1.DName = LCRen Then
        cmdGy.Visible = True
        cmdSave.Enabled = True
        frmJ.Visible = False
    End If
    frmCg.Visible = True
    frmGy.Width = 10005
Exit Sub
End If

frmBan.Visible = True
Call BanQing
cmdSave.Enabled = True
If Me.Lc = 1 And mod1.DName = LCRen Then
    cmdDel.Enabled = True
    frmGui.Enabled = True
    optF.Enabled = True
    optQ.Enabled = True
    frmDj.Enabled = True
End If
If (mod1.DName = "��ı��" Or mod1.DName = "�����" Or mod1.Bm = "�ܾ���" Or mod1.Bm = "������" Or mod1.Qy = "����") And mod1.DName = LCRen And Lc > 1 Or mod1.DName = "������" Then
    frmDj.Enabled = True
End If
If mod1.DName = "��ı��" Or mod1.DName = "�����" Then
    frmDj.Enabled = True

End If
If mod1.DName = "������" Then
    cmdDel.Enabled = True
End If
End Sub

Private Sub cmdNDel_Click()
Dim Did As Long
Did = Val(lblDid.Caption)
If Did = 0 Then Exit Sub

ii = MsgBox("�Ƿ�ɾ���������", vbQuestion + vbYesNo)
If ii = vbNo Then Exit Sub

On Error Resume Next
timZm = 3 '����
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�ɱ�׷�ӵ�"
    mod1.cmd.Parameters("@NBLX") = "����ϸ"
    mod1.cmd.Parameters("@bh") = lblZid.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtSL.Text)  '����
    mod1.cmd.Parameters("@mm2") = Val(txtDj.Text)  '����
    mod1.cmd.Parameters("@mm3") = Val(txtJdj.Text)  '��׼����
    If optGy1.Value = True Then
        mod1.cmd.Parameters("@mm4") = Val(txtGy1.ToolTipText)
    ElseIf optGy2.Value = True Then
        mod1.cmd.Parameters("@mm4") = Val(txtGy2.ToolTipText)
    ElseIf optGy3.Value = True Then
        mod1.cmd.Parameters("@mm4") = Val(txtGY3.ToolTipText)
    End If
    mod1.cmd.Parameters("@mm5") = Val(txtGy1.ToolTipText)
    mod1.cmd.Parameters("@mm6") = Val(txtGy2.ToolTipText)
    mod1.cmd.Parameters("@mm7") = Val(txtGY3.ToolTipText)
    mod1.cmd.Parameters("@mm8") = Val(txtGdj1.Text)
    mod1.cmd.Parameters("@mm9") = Val(txtGdj2.Text)
    mod1.cmd.Parameters("@mm10") = Val(txtGdj3.Text)
    mod1.cmd.Parameters("@mm11") = Did
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 0   '''''''�Ƿ�ɾ��
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

Private Sub cmdNGx_Click()
Dim Did As Long
Did = Val(lblDid.Caption)
If Did = 0 Then Exit Sub

If Val(lblZid.ToolTipText) = 0 Then
    MsgBox "���ȱ���!"
    cmdSave.Enabled = True
    Exit Sub
End If
If Val(txtSL.Text) = 0 Then
    MsgBox "��ȷ������!"
    Exit Sub
End If
On Error Resume Next

timZm = 3 '����
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�ɱ�׷�ӵ�"
    mod1.cmd.Parameters("@NBLX") = "����ϸ"
    mod1.cmd.Parameters("@bh") = lblZid.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtSL.Text)  '����
    mod1.cmd.Parameters("@mm2") = Val(txtDj.Text)  '����
    mod1.cmd.Parameters("@mm3") = Val(txtJdj.Text)  '��׼����
    If optGy1.Value = True Then
        mod1.cmd.Parameters("@mm4") = txtGy1.ToolTipText
    ElseIf optGy2.Value = True Then
        mod1.cmd.Parameters("@mm4") = txtGy2.ToolTipText
    ElseIf optGy3.Value = True Then
        mod1.cmd.Parameters("@mm4") = txtGY3.ToolTipText
    End If
    mod1.cmd.Parameters("@mm5") = Val(txtGy1.ToolTipText)
    mod1.cmd.Parameters("@mm6") = Val(txtGy2.ToolTipText)
    mod1.cmd.Parameters("@mm7") = Val(txtGY3.ToolTipText)
    mod1.cmd.Parameters("@mm8") = Val(txtGdj1.Text)
    mod1.cmd.Parameters("@mm9") = Val(txtGdj2.Text)
    mod1.cmd.Parameters("@mm10") = Val(txtGdj3.Text)
    mod1.cmd.Parameters("@mm11") = Did
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 1   '''''''�Ƿ�ɾ��
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

Private Sub cmdNQ_Click()
Dim tt As String
Dim oo As Integer

Dim ii As Integer


On Error Resume Next




If Lc = 100 Then
    Exit Sub
End If
If LCRen = "�����" And mod1.DName = "��ı��" Then
    LCRen = "��ı��": LCUid = "HM538"
End If
If LCRen <> mod1.DName Then
    MsgBox "�˴�Ӧ��" & lblLcRen.Caption & "ǩ��! ������Ҫ�ٵ�"
    Exit Sub
End If
If Lc = 100 Then

        Exit Sub

End If
If cmdSave.Enabled = True Then
    MsgBox "���Ƚ����ӱ���,��ǩ�����Ĵ���!"
    Exit Sub
End If

    frmQm.Visible = True
    cmdDing.Enabled = True
    
    If Me.Lc = 1 Then   '������ֻ��ǩ�֣����ܲ��ء�
        optT2.Enabled = False
        OptT1.Value = True
    Else
        optT2.Enabled = True
        OptT1.Value = False
        optT2.Value = False
    End If

End Sub

Private Sub cmdSave_Click()
If optQ.Value = False And optF.Value = False Then
    MsgBox "��ȷ���Ƿ�ȫ��!"
    Exit Sub
End If
If comGui.Text = "���ջ���(��˾����)" And optQ.Value = True Then
    MsgBox "�ϰ���ͬû�з��ջ���,������ȷ��!"
    Exit Sub
End If
If comGui.Text = "" Then
    MsgBox "��ȷ�ϳɱ��Ĺ���!"
    Exit Sub
End If


timZm = 1 '����
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�ɱ�׷�ӵ�"
    mod1.cmd.Parameters("@NBLX") = "����"
    mod1.cmd.Parameters("@bh") = Trim(lblBh.ToolTipText)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(lblGLBH.Caption) '��ͬ���
    mod1.cmd.Parameters("@mt2") = comGui.Text
    mod1.cmd.Parameters("@mt3") = lblFBF.Caption
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""

    mod1.cmd.Parameters("@mlt1") = txtBz.Text '��ע
    mod1.cmd.Parameters("@mm1") = Val(lblGLBH.ToolTipText)  'Hid
    mod1.cmd.Parameters("@mm2") = htRow
    If optQ.Value = True Then
        mod1.cmd.Parameters("@mb1") = 1 'ȫ����
    Else
        mod1.cmd.Parameters("@mb1") = 0 'ȫ����
    End If
    If lblFBF.Caption = "���" Or lblFBF.Caption = "����" Or htRow = 9 Or htRow = 10 Or htRow = 11 Or htRow = 12 Or htRow = 13 Or htRow = 14 Or htRow = 15 Or htRow = 16 Or htRow = 17 Or htRow = 18 Or htRow = 19 Then
        mod1.cmd.Parameters("@mb2") = False
    ElseIf InStr(1, lblFBF.Caption, "�ְ�") > 0 Or htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 Or htRow = 8 Or htRow >= 20 Then
        mod1.cmd.Parameters("@mb2") = True
    Else
        MsgBox ("��ȷ�ϴ�׷�ӵ�����������Ƿְ�!")
        Exit Sub
    End If
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

Private Sub comGui_Click()
Dim ii As Integer
If comGui.Text = "" Then Exit Sub
''''''''lblFBF.Caption = ""
''''''''If comGui.Text = "ҵ�񲿷���" Then
''''''''
''''''''ElseIf comGui.Text = "�ɹ�������" Then
''''''''    lblFBF.Caption = "���"
''''''''ElseIf comGui.Text = "���񲿷���" Then
''''''''    lblFBF.Caption = "�ְ�"
''''''''ElseIf comGui.Text = "���ջ���(��˾����)" Then
''''''''
''''''''End If


If lblFBF.Caption = "" Then
    ii = MsgBox("��ȷ����׷�����(Y)?,����׷�ӷְ�(N)?", vbYesNo + vbDefaultButton1 + vbInformation, "��ע��ȷ��!")
    If ii = vbYes Then
        lblFBF.Caption = "���"
    ElseIf ii = vbNo Then
        lblFBF.Caption = "�ְ�"
    End If
End If

End Sub


Private Sub dtgGy_DblClick()
On Error Resume Next
If dtgGy.Row = 0 Then Exit Sub
If GyId = 0 Then GyId = 1
If GyId = 1 Then
    dtgGy.Col = 0: txtGy1.Text = dtgGy.Text
    dtgGy.Col = 1: txtGy1.ToolTipText = dtgGy.Text
    txtGdj1.Text = ""
ElseIf GyId = 2 Then
    dtgGy.Col = 0: txtGy2.Text = dtgGy.Text
    dtgGy.Col = 1: txtGy2.ToolTipText = dtgGy.Text
    txtGdj2.Text = ""
ElseIf GyId = 3 Then
    dtgGy.Col = 0: txtGY3.Text = dtgGy.Text
    dtgGy.Col = 1: txtGY3.ToolTipText = dtgGy.Text
    txtGdj3.Text = ""
End If
End Sub

Private Sub dtgMx_Click()
On Error Resume Next
Call MXQing
dtgN.Row = dtgMx.Row
dtgN.Col = 0: txt1.Text = dtgN.Text
dtgN.Col = 1: txt2.Text = dtgN.Text
dtgN.Col = 2: txt3.Text = dtgN.Text: txtDj.Text = dtgN.Text
dtgN.Col = 3: txt4.Text = dtgN.Text: txtJdj.Text = dtgN.Text
dtgN.Col = 4: txt5.Text = dtgN.Text: txtSL.Text = dtgN.Text
dtgN.Col = 5: txt6.Text = dtgN.Text
dtgN.Col = 6: txt7.Text = dtgN.Text
dtgN.Col = 7: txt7.ToolTipText = dtgN.Text 'Did
lblDid.Caption = dtgN.Text
dtgN.Col = 8: txtGy1.ToolTipText = dtgN.Text
dtgN.Col = 9: txtGy2.ToolTipText = dtgN.Text
dtgN.Col = 10: txtGY3.ToolTipText = dtgN.Text
dtgN.Col = 11: txtGdj1.Text = dtgN.Text
dtgN.Col = 12: txtGdj2.Text = dtgN.Text
dtgN.Col = 13: txtGdj3.Text = dtgN.Text
dtgN.Col = 14: txtGy1.Text = dtgN.Text
dtgN.Col = 15: txtGy2.Text = dtgN.Text
dtgN.Col = 16: txtGY3.Text = dtgN.Text
dtgN.Col = 17
    optGy1.ForeColor = &H80000008: txtGy1.ForeColor = &H80000008: txtGdj1.ForeColor = &H80000008
    optGy2.ForeColor = &H80000008: txtGy2.ForeColor = &H80000008: txtGdj2.ForeColor = &H80000008
    optGy3.ForeColor = &H80000008: txtGY3.ForeColor = &H80000008: txtGdj3.ForeColor = &H80000008
If txtGy1.ToolTipText = dtgN.Text Then
    optGy1.Value = True: optGy1.ForeColor = &HC00000: txtGy1.ForeColor = &HC00000: txtGdj1.ForeColor = &HC00000
ElseIf txtGy2.ToolTipText = dtgN.Text Then
    optGy2.Value = True: optGy2.ForeColor = &HC00000: txtGy2.ForeColor = &HC00000: txtGdj2.ForeColor = &HC00000
ElseIf txtGY3.ToolTipText = dtgN.Text Then
    optGy3.Value = True: optGy3.ForeColor = &HC00000: txtGY3.ForeColor = &HC00000: txtGdj3.ForeColor = &HC00000
End If

If frmCg.Visible = False Then
    frmGy.Width = 6165
End If

End Sub

Private Sub Form_Click()
frmQm.Visible = False
frmCg.Visible = False
frmGy.Visible = False
End Sub

Private Sub Form_DblClick()
Dim ii As Integer
Dim tt As String

Dim Bid1 As Long, Bid6 As Long, Bid7 As Long
Dim Ra
'If mod1.DName <> "���ӱ" Then Exit Sub
If mod1.DName <> "������" Then Exit Sub

Exit Sub

    timZm = 19 'ִ��֪ͨ
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�ɱ�׷�ӵ�"
    mod1.cmd.Parameters("@NBLX") = "ִ��֪ͨ"
    mod1.cmd.Parameters("@bh") = lblZid.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblFBF.Caption '׷������
    mod1.cmd.Parameters("@mt2") = lblKhmc.Caption
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = 0

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

        
    End If

    
Set mod1.cmd = Nothing

Exit Sub
ZXERR:
MsgBox "����!"
End Sub

Private Sub Form_Load()

Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
Call dtgGYFF
frmCg.Top = 5310
frmGy.Top = 5310
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
 dtgP.ColWidth(3) = 6630: dtgP.ColWidth(4) = 1035
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub
Public Sub QMBound(Zid As Long)
Dim Ra: Dim La
Dim ii As Integer: Dim oo As Integer
Dim tt As String
On Error Resume Next

tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & Zid & "' and yid=90 order by pid desc"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2): dtgP.Rows = La + 20
Call dtgPFF
For oo = 1 To La + 1
    dtgP.Row = oo
    For ii = 0 To 5
        dtgP.Col = ii
        dtgP.Text = Ra(ii, oo - 1)
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
For oo = 1 To La + 1
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



End Sub

Public Sub Qing()
lblKhmc.Caption = ""
lblGLBH.Caption = ""
lblGLBH.ToolTipText = ""
Me.optQ.Value = False
Me.optF.Value = False
lblZbh.Caption = ""
lblXz.Caption = ""
lblZe.Caption = ""
txtBz.Text = ""
dtgMx.Clear
comGui.Text = ""
lblBh.Caption = ""
lblBh.ToolTipText = ""

Call BanQing
frmBan.Visible = False
LCRen = ""
LCUid = ""
Lc = 0
Fwid = 0
 NewF = 0
 Hid = 0
lblZid.ToolTipText = ""
lblJe.Caption = ""
cmdSave.Enabled = False
cmdDel.Enabled = False
frmDj.Enabled = False
lblZtime.Caption = ""

lblYwy.Caption = ""
lblYwy.ToolTipText = ""

frmGui.Visible = False
lblFBF.Caption = ""
txtQM.Text = ""
optF.Enabled = False
optQ.Enabled = False
Me.frmGui.Enabled = False
lblCb1.Caption = ""
lblCB2.Caption = ""
lblMF.Caption = ""

frmCg.Visible = False
frmGy.Visible = False
cmdGy.Visible = False
End Sub

Public Sub dtgFF()
dtgMx.Clear
dtgMx.Cols = 8
dtgMx.Rows = 20
dtgMx.Row = 0

dtgMx.Col = 0: dtgMx.Text = "��/������/��������": dtgMx.CellFontBold = True
dtgMx.Col = 1: dtgMx.Text = "����Ʒ��(�ͺ�)": dtgMx.CellFontBold = True
dtgMx.Col = 2: dtgMx.Text = "����ͺ�": dtgMx.CellFontBold = True
dtgMx.Col = 3: dtgMx.Text = "��Ӧ/�ְ���": dtgMx.CellFontBold = True
dtgMx.Col = 4: dtgMx.Text = "����": dtgMx.CellFontBold = True
dtgMx.Col = 5: dtgMx.Text = "����": dtgMx.CellFontBold = True
dtgMx.Col = 6: dtgMx.Text = "�ܼ�": dtgMx.CellFontBold = True
dtgMx.ColWidth(7) = 0
dtgMx.ColWidth(0) = 5115
dtgMx.ColWidth(1) = 2265
dtgMx.ColWidth(2) = 1725
dtgMx.ColWidth(3) = 2100

dtgN.Clear
dtgN.Cols = 8
dtgN.Rows = 20
dtgN.Row = 0


End Sub
Public Sub dtgFF1()
dtgMx.Clear
dtgMx.Cols = 18
dtgMx.Rows = 20
dtgMx.Row = 0

dtgMx.Col = 0: dtgMx.Text = "��Ʒ���": dtgMx.CellFontBold = True
dtgMx.Col = 1: dtgMx.Text = "��Ʒ����": dtgMx.CellFontBold = True
dtgMx.Col = 2: dtgMx.Text = "����": dtgMx.CellFontBold = True
dtgMx.Col = 3: dtgMx.Text = "���۵���": dtgMx.CellFontBold = True
dtgMx.Col = 4: dtgMx.Text = "����": dtgMx.CellFontBold = True
dtgMx.Col = 5: dtgMx.Text = "С��": dtgMx.CellFontBold = True
dtgMx.Col = 6: dtgMx.Text = "��Ч��": dtgMx.CellFontBold = True
dtgMx.Col = 7: dtgMx.Text = "did": dtgMx.CellFontBold = True
dtgMx.Col = 8: dtgMx.Text = "gyid1": dtgMx.CellFontBold = True
dtgMx.Col = 9: dtgMx.Text = "gyid2": dtgMx.CellFontBold = True
dtgMx.Col = 10: dtgMx.Text = "gyid3": dtgMx.CellFontBold = True
dtgMx.Col = 11: dtgMx.Text = "gdj1": dtgMx.CellFontBold = True
dtgMx.Col = 12: dtgMx.Text = "gdj2": dtgMx.CellFontBold = True
dtgMx.Col = 13: dtgMx.Text = "gdj3": dtgMx.CellFontBold = True
dtgMx.Col = 14: dtgMx.Text = "mc1": dtgMx.CellFontBold = True
dtgMx.Col = 15: dtgMx.Text = "mc2": dtgMx.CellFontBold = True
dtgMx.Col = 16: dtgMx.Text = "mc3": dtgMx.CellFontBold = True
dtgMx.Col = 17: dtgMx.Text = "gyid": dtgMx.CellFontBold = True


dtgMx.ColWidth(0) = -1
dtgMx.ColWidth(1) = 9225
dtgMx.ColWidth(2) = -1
dtgMx.ColWidth(3) = -1
dtgMx.ColWidth(4) = -1
dtgMx.ColWidth(5) = -1
dtgMx.ColWidth(7) = 0
dtgMx.ColWidth(8) = 0
dtgMx.ColWidth(9) = 0
dtgMx.ColWidth(10) = 0
dtgMx.ColWidth(11) = 0
dtgMx.ColWidth(12) = 0
dtgMx.ColWidth(13) = 0
dtgMx.ColWidth(14) = 0
dtgMx.ColWidth(15) = 0
dtgMx.ColWidth(16) = 0
dtgMx.ColWidth(17) = 0
dtgN.Clear
dtgN.Cols = 18
dtgN.Rows = 20
dtgN.Row = 0


End Sub
Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
Cancel = True
End Sub

Private Sub Label24_DblClick()
If mod1.DName = "������" Or mod1.DName = "������" Or mod1.DName = "�Ǽ���" Or mod1.DName = "��ȫ��" Then
    frmJ.Visible = False
End If
End Sub





Public Sub dtgGYFF()
dtgGy.Clear
dtgGy.Rows = 50
dtgGy.Cols = 2
dtgGy.Row = 0
dtgGy.Col = 0: dtgGy.Text = "��Ӧ�����ƣ����˫��ѡ��": dtgGy.CellFontBold = True
dtgGy.ColWidth(1) = 0
dtgGy.ColWidth(0) = 3480

End Sub

Private Sub optGy1_Click()
optGy1.ForeColor = &H80000008: txtGy1.ForeColor = &H80000008: txtGdj1.ForeColor = &H80000008
optGy2.ForeColor = &H80000008: txtGy2.ForeColor = &H80000008: txtGdj2.ForeColor = &H80000008
optGy3.ForeColor = &H80000008: txtGY3.ForeColor = &H80000008: txtGdj3.ForeColor = &H80000008
If optGy1.Value = True Then
    optGy1.ForeColor = &HC00000: txtGy1.ForeColor = &HC00000: txtGdj1.ForeColor = &HC00000
    txtDj.Text = txtGdj1.Text
    If Val(lblMF.Caption) > 0.55 Then
        txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * Val(lblMF.Caption), 2)
    Else
        txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * 0.55, 2)
    End If
End If
End Sub

Private Sub optGy2_Click()
optGy1.ForeColor = &H80000008: txtGy1.ForeColor = &H80000008: txtGdj1.ForeColor = &H80000008
optGy2.ForeColor = &H80000008: txtGy2.ForeColor = &H80000008: txtGdj2.ForeColor = &H80000008
optGy3.ForeColor = &H80000008: txtGY3.ForeColor = &H80000008: txtGdj3.ForeColor = &H80000008
If optGy2.Value = True Then
    optGy2.ForeColor = &HC00000: txtGy2.ForeColor = &HC00000: txtGdj2.ForeColor = &HC00000
    txtDj.Text = txtGdj2.Text
    If Val(lblMF.Caption) > 0.55 Then
        txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * Val(lblMF.Caption), 2)
    Else
        txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * 0.55, 2)
    End If
End If
End Sub


Private Sub optGy3_Click()
optGy1.ForeColor = &H80000008: txtGy1.ForeColor = &H80000008: txtGdj1.ForeColor = &H80000008
optGy2.ForeColor = &H80000008: txtGy2.ForeColor = &H80000008: txtGdj2.ForeColor = &H80000008
optGy3.ForeColor = &H80000008: txtGY3.ForeColor = &H80000008: txtGdj3.ForeColor = &H80000008
If optGy3.Value = True Then
    optGy3.ForeColor = &HC00000: txtGY3.ForeColor = &HC00000: txtGdj3.ForeColor = &HC00000
    txtDj.Text = txtGdj3.Text
    If Val(lblMF.Caption) > 0.55 Then
        txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * Val(lblMF.Caption), 2)
    Else
        txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * 0.55, 2)
    End If

End If
End Sub


Private Sub timQuit_Timer()
Dim oo As Integer
Dim ii As Integer
Dim Rf
Dim Rg
On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
Dim tt As String
If timZm = 1 Then '���Ϊ��Ӻ�ͬ����
    cmdSave.Enabled = False
    Call FmxcNew.LXBound(Rf, Rg)
ElseIf timZm = 2 Then
    Me.Visible = False 'ɾ��
    If FmxcNew.Visible = True Then
        Call FmxcNew.LXBound(Rf, Rg)
    End If
ElseIf timZm = 5 Then 'ǩ��
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
    If Dialog.Visible = True Then
    Call mod1.refEnvent(1)
    End If
ElseIf timZm = 19 Then 'ִ��֪ͨ
    MsgBox "�Ѿ��ɹ�֪ͨ:" & lblTX.Caption & "!"
End If
timQuit.Enabled = False

End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
Dim RC, RD, RE
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2,mt3 from ml where zid=" & mod1.Zid
Set mod1.WP = New ADODB.Recordset
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '�ύ�ɹ�
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdSave.Enabled = False
        lblZid.ToolTipText = mod1.WP.Fields("mm1").Value
        lblBh.ToolTipText = mod1.WP.Fields("mt1").Value
        lblBh.Caption = Right(lblBh.ToolTipText, 3)
        If Left(lblBh.Caption, 1) = "J" Then
            lblBh.Caption = Right(lblBh.ToolTipText, 4)
        End If
    ElseIf timZm = 3 Then
    tt = "declare @hid int;" & _
        "select @hid=hid from htzui where zid=" & Zid & ";" & _
        "select bh,nr,dj,jdj,sl,ze,delf,did,gyid1,gyid2,gyid3,gdj1,gdj2,gdj3,mc1,mc2,mc3,gyid from zuijiaDetail where zid=" & Val(lblZid.ToolTipText) & " order by delf desc,did;" & _
        "select sum(ze) from htzuidetail where zid=" & Val(lblZid.ToolTipText) & ";" & _
        "select sum(ze) from htzuiZe where hid=@hid"
        Set mod1.HTP = New ADODB.Recordset
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        On Error Resume Next
        RC = mod1.HTP.GetRows
        Set mod1.HTP = mod1.HTP.NextRecordset
        RD = mod1.HTP.GetRows
        Set mod1.HTP = mod1.HTP.NextRecordset
        RE = mod1.HTP.GetRows
        Set mod1.HTP = Nothing
        Call Me.NewMxBound(RC, RD, RE)
    ElseIf timZm = 6 Then
    tt = "declare @hid int;" & _
        "select @hid=hid from htzui where zid=" & Zid & ";" & _
        "select bh,nr,dj,jdj,sl,ze,delf,did,gyid1,gyid2,gyid3,gdj1,gdj2,gdj3,mc1,mc2,mc3,gyid from zuijiaDetail where zid=" & Val(lblZid.ToolTipText) & " order by delf desc,did;" & _
        "select sum(ze) from htzuidetail where zid=" & Val(lblZid.ToolTipText) & ";" & _
        "select sum(ze) from htzuiZe where hid=@hid"
        Set mod1.HTP = New ADODB.Recordset
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        On Error Resume Next
        RC = mod1.HTP.GetRows
        Set mod1.HTP = mod1.HTP.NextRecordset
        RD = mod1.HTP.GetRows
        Set mod1.HTP = mod1.HTP.NextRecordset
        RE = mod1.HTP.GetRows
        Set mod1.HTP = Nothing
        Call Me.NewMxBound(RC, RD, RE)
    ElseIf timZm = 5 Then
        frmQm.Visible = False
        Me.Lc = mod1.WP.Fields("mm1").Value
        Fwid = mod1.WP.Fields("mm2").Value
        LCRen = mod1.WP.Fields("mt1").Value
        LCUid = mod1.WP.Fields("mt2").Value
        lblTX.Caption = "��ǰ������:" & LCRen
        If Me.Lc = 100 Then
            lblTX.Caption = "���̽���"
        End If
        
        Call QMBound(Val(lblZid.ToolTipText))
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
    If timZm = 1 Then
        cmdSave.Enabled = False
    End If
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("���������ڴ�����������ʱ,��ʱ!", vbExclamation + vbOKOnly, "��������!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdSave.Enabled = False
    End If
    Exit Sub

End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


Public Sub BanQing()
txt1.Text = ""
txt2.Text = ""
txt3.Text = ""
txt4.Text = ""
txt5.Text = ""
txt6.Text = ""
txt7.Text = ""
txt7.ToolTipText = ""
End Sub

Public Sub Bound(Zid As Long)
Dim NewFZ As Integer
Dim tt As String
Dim Ra, Rb, RC, RD, RE, Rf
Dim La
Dim Lc As Integer
Dim oo As Integer
Dim QBZE As Single
Dim Yj As Single
Dim CBZE As Single
Call Me.Qing
tt = "select newF from htzui where zid=" & Zid
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
NewFZJ = Ra(0, 0): Set Ra = Nothing
If NewFZJ = 888 Then
    tt = "declare @hid int;" & _
        "select @hid=hid from htzui where zid=" & Zid & ";" & _
        "select htbh,qbf,bh,bz,gui,hid,lcren,lcuid,lc,fwid,zid,ywy,uid,XZ,fbf,ztime,zl,htrow from htzui where zid=" & Zid & ";" & _
        "select khmc,zbh,htxz,htze,clf,newF,hid,qbze,yj  from htping where hid=@hid;" & _
        "select nr,pb,xh,gyfb,sl,dj,ze,did from htzuidetail where zid=" & Zid & " order by did;" & _
        "select sum(ze) from htzuidetail where zid=" & Zid & " and delf=1;" & _
        "select sum(ze) from htzuiZe where hid=@hid"
ElseIf NewFZJ = 1 Or NewFZJ = 0 Then
    tt = "declare @hid int;" & _
        "select @hid=hid from htzui where zid=" & Zid & ";" & _
        "select htbh,qbf,bh,bz,gui,hid,lcren,lcuid,lc,fwid,zid,ywy,uid,XZ,fbf,ztime,zl,htrow from htzui where zid=" & Zid & ";" & _
        "select khmc,zbh,htxz,htze,clf,newF,hid,qbze,yj from htping where hid=@hid;" & _
        "select bh,nr,dj,jdj,sl,ze,delf,did,gyid1,gyid2,gyid3,gdj1,gdj2,gdj3,mc1,mc2,mc3,gyid from zuijiaDetail where zid=" & Zid & " and delf=1 order by delf desc,did;" & _
        "select sum(ze) from htzuidetail where zid=" & Zid & " and delf=1;" & _
        "select sum(ze) from htzuiZe where hid=@hid;" & _
        "select sum(Jhg) from xunjiaD where cast(htbh as int)=@hid and delf=1 and lc=100 and not(zl like '%Ԥ��%')"
End If

Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RD = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RE = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rf = mod1.HTP.GetRows
Set mod1.HTP = Nothing


Lc = UBound(RC, 2) + 1
lblKhmc.Caption = Rb(0, 0)
lblGLBH.Caption = Ra(0, 0)
lblGLBH.ToolTipText = Ra(5, 0)

If Ra(1, 0) = True Then
    Me.optQ.Value = True
ElseIf Ra(1, 0) = False Then
    Me.optF.Value = True
End If
lblZbh.Caption = Rb(1, 0)
lblXz.Caption = Rb(2, 0)
lblZe.Caption = Rb(3, 0)
lblCb1.Caption = Rb(4, 0)
NewF = Rb(5, 0)
Hid = Rb(6, 0)
QBZE = Rb(7, 0)
Yj = Rb(8, 0)

txtBz.Text = Ra(3, 0)
comGui.Text = Ra(4, 0)
lblBh.ToolTipText = Ra(2, 0)
lblBh.Caption = Right(lblBh.ToolTipText, 3)
If Left(lblBh.Caption, 1) = "J" Then
    lblBh.Caption = Right(lblBh.ToolTipText, 4)
End If
LCRen = Ra(6, 0)
LCUid = Ra(7, 0)
Me.Lc = Ra(8, 0)
Fwid = Ra(9, 0)
lblZid.ToolTipText = Ra(10, 0)
lblYwy.Caption = Ra(11, 0)
lblYwy.ToolTipText = Ra(12, 0)

lblCB2.Caption = RE(0, 0) '�Ѿ���������ӳɱ�
xZ = Ra(13, 0)
fmxcZJ.htRow = Ra(17, 0)
lblFBF.Caption = Ra(16, 0)
If lblFBF.Caption = "" Then '�ɰ汾���ʵ���ʾ
    If Ra(14, 0) = True Then
        lblFBF.Caption = "�ְ�"
    Else
        lblFBF.Caption = "���"
    End If
End If
If IsNull(Ra(15, 0)) = False Then
    lblZtime.Caption = "ִ��ʱ�� " & Ra(15, 0)
End If

lblJe.Caption = RD(0, 0)
lblTX.Caption = "����������:" & LCRen: lblTX.Visible = True
        If Me.Lc = 100 Then
            lblTX.Caption = "���̽���"
        ElseIf Me.Lc = 101 Then
            lblTX.Caption = "ִ�н׶�"
        End If
If NewFZJ = 888 Then
    Call dtgFF
    For oo = 1 To Lc
        dtgMx.Row = oo
        dtgMx.Col = 0: dtgMx.Text = RC(0, oo - 1)
        dtgMx.Col = 1: dtgMx.Text = RC(1, oo - 1)
        dtgMx.Col = 2: dtgMx.Text = RC(2, oo - 1)
        dtgMx.Col = 3: dtgMx.Text = RC(3, oo - 1)
        dtgMx.Col = 4: dtgMx.Text = RC(4, oo - 1)
        dtgMx.Col = 5: dtgMx.Text = RC(5, oo - 1)
        dtgMx.Col = 6: dtgMx.Text = RC(6, oo - 1)
        dtgMx.Col = 7: dtgMx.Text = RC(7, oo - 1)
        dtgN.Row = oo
        dtgN.Col = 0: dtgN.Text = RC(0, oo - 1)
        dtgN.Col = 1: dtgN.Text = RC(1, oo - 1)
        dtgN.Col = 2: dtgN.Text = RC(2, oo - 1)
        dtgN.Col = 3: dtgN.Text = RC(3, oo - 1)
        dtgN.Col = 4: dtgN.Text = RC(4, oo - 1)
        dtgN.Col = 5: dtgN.Text = RC(5, oo - 1)
        dtgN.Col = 6: dtgN.Text = RC(6, oo - 1)
        dtgN.Col = 7: dtgN.Text = RC(7, oo - 1)
    Next
    Call BanQing
ElseIf NewFZJ = 1 Or NewFZJ = 0 Then
    Call Me.NewMxBound(RC, RD, RE)
End If
Call QMBound(Val(lblZid.ToolTipText))
frmBan.Visible = False
If xZ = False Then
    frmGui.Visible = True
    Me.Caption = "�ɱ�׷�ӵ�"
Else
    frmGui.Visible = False
    Me.Caption = "���˵��"
End If
CBZE = Rf(0, 0)
lblMF.Caption = Round((Val(lblZe.Caption) - Yj - QBZE) / CBZE, 2)
'Me.lblMF.Caption = Round((Val(txtHtze.Text) - Val(txtYJ.Text) - QBZE) / Val(lblCBZE.Caption), 2)
End Sub

Private Sub txt5_Change()
txt7.Text = Val(txt5.Text) * Val(txt6.Text)

End Sub



Public Sub MXBound(Zid As Long)
Dim tt As String
Dim RC, RD, RE
Dim Lc As Integer
Dim oo As Integer
'''''If NewFZJ = 0 Then
    tt = "select nr,pb,xh,gyfb,sl,dj,ze,did from htzuidetail where zid=" & Zid & " order by did;" & _
        "select sum(ze) from htzuidetail where zid=" & Zid & " and delf=1;" & _
        "select sum(ze) from htzuiZe where hid=" & Hid
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    On Error Resume Next
    RC = mod1.HTP.GetRows
    Set mod1.HTP = mod1.HTP.NextRecordset
    RD = mod1.HTP.GetRows
    Set mod1.HTP = mod1.HTP.NextRecordset
    RE = mod1.HTP.GetRows
    Set mod1.HTP = Nothing
    Lc = UBound(RC, 2) + 1
    Call dtgFF
    For oo = 1 To Lc
        dtgMx.Row = oo
        dtgMx.Col = 0: dtgMx.Text = RC(0, oo - 1)
        dtgMx.Col = 1: dtgMx.Text = RC(1, oo - 1)
        dtgMx.Col = 2: dtgMx.Text = RC(2, oo - 1)
        dtgMx.Col = 3: dtgMx.Text = RC(3, oo - 1)
        dtgMx.Col = 4: dtgMx.Text = RC(4, oo - 1)
        dtgMx.Col = 5: dtgMx.Text = RC(5, oo - 1)
        dtgMx.Col = 6: dtgMx.Text = RC(6, oo - 1)
        dtgMx.Col = 7: dtgMx.Text = RC(7, oo - 1)
        dtgN.Row = oo
        dtgN.Col = 0: dtgN.Text = RC(0, oo - 1)
        dtgN.Col = 1: dtgN.Text = RC(1, oo - 1)
        dtgN.Col = 2: dtgN.Text = RC(2, oo - 1)
        dtgN.Col = 3: dtgN.Text = RC(3, oo - 1)
        dtgN.Col = 4: dtgN.Text = RC(4, oo - 1)
        dtgN.Col = 5: dtgN.Text = RC(5, oo - 1)
        dtgN.Col = 6: dtgN.Text = RC(6, oo - 1)
        dtgN.Col = 7: dtgN.Text = RC(7, oo - 1)
    Next
    lblJe.Caption = RD(0, 0)
    lblCB2.Caption = RE(0, 0)
'''''ElseIf NewFZJ = 1 Then
'''''    tt = "select nr,pb,xh,gyfb,sl,dj,ze,did from htzuidetail where zid=" & Zid & " order by did;" & _
'''''        "select sum(ze) from htzuidetail where zid=" & Zid & ";" & _
'''''        "select sum(ze) from htzuiZe where hid=" & Hid
'''''    Set mod1.HTP = New ADODB.Recordset
'''''    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''    On Error Resume Next
'''''    RC = mod1.HTP.GetRows
'''''    Set mod1.HTP = mod1.HTP.NextRecordset
'''''    RD = mod1.HTP.GetRows
'''''    Set mod1.HTP = mod1.HTP.NextRecordset
'''''    RE = mod1.HTP.GetRows
'''''    Set mod1.HTP = Nothing
'''''    Lc = UBound(RC, 2) + 1
'''''    Call dtgFF
'''''    For oo = 1 To Lc
'''''        dtgMx.Row = oo
'''''        dtgMx.Col = 0: dtgMx.Text = RC(0, oo - 1)
'''''        dtgMx.Col = 1: dtgMx.Text = RC(1, oo - 1)
'''''        dtgMx.Col = 2: dtgMx.Text = RC(2, oo - 1)
'''''        dtgMx.Col = 3: dtgMx.Text = RC(3, oo - 1)
'''''        dtgMx.Col = 4: dtgMx.Text = RC(4, oo - 1)
'''''        dtgMx.Col = 5: dtgMx.Text = RC(5, oo - 1)
'''''        dtgMx.Col = 6: dtgMx.Text = RC(6, oo - 1)
'''''        dtgMx.Col = 7: dtgMx.Text = RC(7, oo - 1)
'''''        dtgN.Row = oo
'''''        dtgN.Col = 0: dtgN.Text = RC(0, oo - 1)
'''''        dtgN.Col = 1: dtgN.Text = RC(1, oo - 1)
'''''        dtgN.Col = 2: dtgN.Text = RC(2, oo - 1)
'''''        dtgN.Col = 3: dtgN.Text = RC(3, oo - 1)
'''''        dtgN.Col = 4: dtgN.Text = RC(4, oo - 1)
'''''        dtgN.Col = 5: dtgN.Text = RC(5, oo - 1)
'''''        dtgN.Col = 6: dtgN.Text = RC(6, oo - 1)
'''''        dtgN.Col = 7: dtgN.Text = RC(7, oo - 1)
'''''    Next
'''''    lblJe.Caption = RD(0, 0)
'''''    lblCB2.Caption = RE(0, 0)
'''''End If
End Sub

Private Sub txt6_Change()
txt7.Text = Val(txt5.Text) * Val(txt6.Text)
End Sub

Private Sub txt7_Change()
If Me.Visible = False Then Exit Sub
On Error Resume Next
txt6.Text = Round(Val(txt7.Text) / Val(txt5.Text), 2)
End Sub

Public Sub NewMxBound(RC, RD, RE)
Dim Lc As Integer
On Error Resume Next
    Lc = UBound(RC, 2) + 1
    
    Call dtgFF1
    dtgMx.Rows = Lc + 50: dtgN.Rows = Lc + 50
    On Error Resume Next
    For oo = 1 To Lc
        dtgMx.Row = oo: dtgMx.RowHeight(oo) = dtgMx.RowHeight(0) * 2
        dtgMx.Col = 0: dtgMx.Text = RC(0, oo - 1)
        dtgMx.Col = 1: dtgMx.Text = RC(1, oo - 1)
        dtgMx.Col = 2: dtgMx.Text = RC(2, oo - 1)
        dtgMx.Col = 3: dtgMx.Text = RC(3, oo - 1)
        dtgMx.Col = 4: dtgMx.Text = RC(4, oo - 1)
        dtgMx.Col = 5: dtgMx.Text = RC(5, oo - 1)
        dtgMx.Col = 6: dtgMx.Text = RC(6, oo - 1)
        dtgMx.Col = 7: dtgMx.Text = RC(7, oo - 1)
        dtgMx.Col = 8: dtgMx.Text = RC(8, oo - 1)
        dtgMx.Col = 9: dtgMx.Text = RC(9, oo - 1)
        dtgMx.Col = 10: dtgMx.Text = RC(10, oo - 1)
        dtgMx.Col = 11: dtgMx.Text = RC(11, oo - 1)
        dtgMx.Col = 12: dtgMx.Text = RC(12, oo - 1)
        dtgMx.Col = 13: dtgMx.Text = RC(13, oo - 1)
        dtgMx.Col = 14: dtgMx.Text = RC(14, oo - 1)
        dtgMx.Col = 15: dtgMx.Text = RC(15, oo - 1)
        dtgMx.Col = 16: dtgMx.Text = RC(16, oo - 1)
        dtgMx.Col = 17: dtgMx.Text = RC(17, oo - 1)
        dtgN.Row = oo
        dtgN.Col = 0: dtgN.Text = RC(0, oo - 1)
        dtgN.Col = 1: dtgN.Text = RC(1, oo - 1)
        dtgN.Col = 2: dtgN.Text = RC(2, oo - 1)
        dtgN.Col = 3: dtgN.Text = RC(3, oo - 1)
        dtgN.Col = 4: dtgN.Text = RC(4, oo - 1)
        dtgN.Col = 5: dtgN.Text = RC(5, oo - 1)
        dtgN.Col = 6: dtgN.Text = RC(6, oo - 1)
        dtgN.Col = 7: dtgN.Text = RC(7, oo - 1)
        dtgN.Col = 8: dtgN.Text = RC(8, oo - 1)
        dtgN.Col = 9: dtgN.Text = RC(9, oo - 1)
        dtgN.Col = 10: dtgN.Text = RC(10, oo - 1)
        dtgN.Col = 11: dtgN.Text = RC(11, oo - 1)
        dtgN.Col = 12: dtgN.Text = RC(12, oo - 1)
        dtgN.Col = 13: dtgN.Text = RC(13, oo - 1)
        dtgN.Col = 14: dtgN.Text = RC(14, oo - 1)
        dtgN.Col = 15: dtgN.Text = RC(15, oo - 1)
        dtgN.Col = 16: dtgN.Text = RC(16, oo - 1)
        dtgN.Col = 17: dtgN.Text = RC(17, oo - 1)
    Next
    lblJe.Caption = RD(0, 0)
    lblCB2.Caption = RE(0, 0)
End Sub

Public Sub MXQing()
txtDj.Text = ""
txtJdj.Text = ""
txtSL.Text = ""
lblDid.Caption = ""
optGy1.Value = False
optGy2.Value = False
optGy3.Value = False
txtGy1.Text = "": txtGy1.ToolTipText = ""
txtGy2.Text = "": txtGy2.ToolTipText = ""
txtGY3.Text = "": txtGY3.ToolTipText = ""
txtGdj1.Text = ""
txtGdj2.Text = ""
txtGdj3.Text = ""
End Sub

Private Sub txtDj_Change()
'txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ, 2)
If Val(lblMF.Caption) > 0.55 Then
    txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * Val(lblMF.Caption), 2)
Else
    txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * 0.55, 2)
End If
End Sub


Private Sub txtGdj1_Change()
If optGy1.Value = True Then
    txtDj.Text = txtGdj1.Text
    If Val(lblMF.Caption) > 0.55 Then
        txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * lblMF.Caption, 2)
    Else
        txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * 0.55, 2)
    End If
End If
End Sub

Private Sub txtGdj2_Change()
If optGy2.Value = True Then
    txtDj.Text = txtGdj2.Text
    If Val(lblMF.Caption) > 0.55 Then
        txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * Val(lblMF.Caption), 2)
    Else
        txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * 0.55, 2)
    End If
End If
End Sub

Private Sub txtGdj3_Change()
If optGy3.Value = True Then
    txtDj.Text = txtGdj3.Text
    If Val(lblMF.Caption) > 0.55 Then
        txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * Val(lblMF.Caption), 2)
    Else
        txtJdj.Text = Round(Val(txtDj.Text) * mod1.JZ * 0.55, 2)
    End If

End If
End Sub


Private Sub txtGy_Change()
Dim tt As String
Dim Ra
Dim La As Long
Dim oo As Long
If Len(txtGy.Text) < 2 Then Exit Sub
'tt = "select mc,gid from gymxc where mc like '%" & txtGy.Text & "%' and delf=1 and lc=100"
tt = "select mc,gid from gymxc where mc like '%" & txtGy.Text & "%' and delf=1"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Call Me.dtgGYFF
For oo = 1 To La
    dtgGy.Row = oo
    dtgGy.Col = 0: dtgGy.Text = Ra(0, oo - 1)
    dtgGy.Col = 1: dtgGy.Text = Ra(1, oo - 1)
Next
End Sub

Private Sub txtGy1_Click()
GyId = 1
End Sub


Private Sub txtGy1_DblClick()
On Error Resume Next
Dim Gid As Long

Gid = Val(txtGy1.ToolTipText)
'If Gid = 0 Then Exit Sub
Call frmGyDetail.Qing
Call frmGyDetail.Bound(Gid)
frmGyDetail.cmdSave.Enabled = False
frmGyDetail.Show
frmGyDetail.ZOrder 0
End Sub


Private Sub txtGy2_Click()
GyId = 2
End Sub


Private Sub txtGy2_DblClick()
On Error Resume Next
Dim Gid As Long

Gid = Val(txtGy2.ToolTipText)
'If Gid = 0 Then Exit Sub
Call frmGyDetail.Qing
Call frmGyDetail.Bound(Gid)
frmGyDetail.cmdSave.Enabled = False
frmGyDetail.Show
frmGyDetail.ZOrder 0
End Sub


Private Sub txtGy3_Click()
GyId = 3
End Sub


Private Sub txtGy3_DblClick()
On Error Resume Next
Dim Gid As Long

Gid = Val(txtGY3.ToolTipText)
'If Gid = 0 Then Exit Sub
Call frmGyDetail.Qing
Call frmGyDetail.Bound(Gid)
frmGyDetail.cmdSave.Enabled = False
frmGyDetail.Show
frmGyDetail.ZOrder 0
End Sub


