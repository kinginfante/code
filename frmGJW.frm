VERSION 5.00
Begin VB.Form frmGJW 
   Caption         =   "施工计划表"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.TextBox txtBid 
      Height          =   285
      Left            =   13440
      TabIndex        =   96
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   4
      Left            =   4560
      TabIndex        =   92
      Top             =   8280
      Width           =   945
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   3
      Left            =   3480
      TabIndex        =   89
      Top             =   8280
      Width           =   945
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   2
      Left            =   2400
      TabIndex        =   86
      Top             =   8280
      Width           =   945
   End
   Begin VB.Frame frmHide 
      Caption         =   "frmHid"
      Height          =   1305
      Left            =   8070
      TabIndex        =   76
      Top             =   7890
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label lblGid 
         Caption         =   "lblGid"
         Height          =   255
         Left            =   3450
         TabIndex        =   85
         Top             =   780
         Width           =   675
      End
      Begin VB.Label lblTrq 
         Caption         =   "lblTrq"
         Height          =   225
         Left            =   3390
         TabIndex        =   84
         Top             =   210
         Width           =   735
      End
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   315
         Left            =   240
         TabIndex        =   83
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lblLcRen 
         Caption         =   "lblLcRen"
         Height          =   285
         Left            =   180
         TabIndex        =   82
         Top             =   570
         Width           =   795
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   180
         TabIndex        =   81
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   1380
         TabIndex        =   80
         Top             =   210
         Width           =   885
      End
      Begin VB.Label lblUid 
         Caption         =   "lblUid"
         Height          =   255
         Left            =   2580
         TabIndex        =   79
         Top             =   780
         Width           =   975
      End
      Begin VB.Label lblYwy 
         Caption         =   "lblYwy"
         Height          =   285
         Left            =   2520
         TabIndex        =   78
         Top             =   450
         Width           =   765
      End
      Begin VB.Label lblPwf 
         Caption         =   "lblPwf"
         Height          =   225
         Left            =   3480
         TabIndex        =   77
         Top             =   450
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   1
      Left            =   1320
      TabIndex        =   73
      Top             =   8280
      Width           =   945
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   20
      Left            =   13650
      TabIndex        =   72
      Top             =   600
      Width           =   585
   End
   Begin VB.CommandButton cmdQm 
      Caption         =   "cmdQm"
      Height          =   345
      Index           =   0
      Left            =   240
      TabIndex        =   69
      Top             =   8280
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   555
      Left            =   13860
      Picture         =   "frmGJW.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   8550
      Width           =   735
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   555
      Left            =   14610
      Picture         =   "frmGJW.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   8550
      Width           =   645
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   555
      Left            =   13140
      Picture         =   "frmGJW.frx":076C
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   8550
      Width           =   705
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   19
      Left            =   13140
      TabIndex        =   65
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   18
      Left            =   12690
      TabIndex        =   64
      Top             =   600
      Width           =   465
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   17
      Left            =   12225
      TabIndex        =   63
      Top             =   600
      Width           =   465
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   16
      Left            =   11790
      TabIndex        =   62
      Top             =   600
      Width           =   435
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   15
      Left            =   11310
      TabIndex        =   61
      Top             =   600
      Width           =   465
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   14
      Left            =   10860
      TabIndex        =   60
      Top             =   600
      Width           =   465
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   13
      Left            =   10395
      TabIndex        =   59
      Top             =   600
      Width           =   465
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   12
      Left            =   9945
      TabIndex        =   58
      Top             =   600
      Width           =   435
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   11
      Left            =   9480
      TabIndex        =   57
      Top             =   600
      Width           =   435
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   10
      Left            =   9030
      TabIndex        =   56
      Top             =   600
      Width           =   435
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   9
      Left            =   8550
      TabIndex        =   55
      Top             =   600
      Width           =   435
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   8
      Left            =   8115
      TabIndex        =   54
      Top             =   600
      Width           =   435
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   7
      Left            =   7665
      TabIndex        =   53
      Top             =   600
      Width           =   435
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   6
      Left            =   7230
      TabIndex        =   52
      Top             =   600
      Width           =   405
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   5
      Left            =   6750
      TabIndex        =   51
      Top             =   600
      Width           =   435
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   4
      Left            =   6270
      TabIndex        =   50
      Top             =   600
      Width           =   435
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   3
      Left            =   5835
      TabIndex        =   49
      Top             =   600
      Width           =   405
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   2
      Left            =   5385
      TabIndex        =   48
      Top             =   600
      Width           =   435
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   1
      Left            =   4926
      TabIndex        =   47
      Top             =   600
      Width           =   405
   End
   Begin VB.TextBox txtDate 
      Height          =   405
      Left            =   810
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Top             =   600
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   14
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   28
      Top             =   6840
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   13
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   6420
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   12
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   5970
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   11
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   5520
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   10
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   5070
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   9
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   4620
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   8
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   4170
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   7
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Top             =   3720
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   6
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   3270
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   5
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   2820
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   4
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   2400
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   3
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   1950
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   2
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   1500
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   1
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   1050
      Width           =   3555
   End
   Begin VB.TextBox txtNr 
      Height          =   435
      Index           =   15
      Left            =   810
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   7290
      Width           =   3555
   End
   Begin VB.TextBox txtZu 
      Height          =   285
      Left            =   10500
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   30
      Width           =   1725
   End
   Begin VB.TextBox txtXMMC 
      Height          =   285
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   30
      Width           =   3615
   End
   Begin VB.TextBox txtHtbh 
      Height          =   270
      Left            =   6150
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   3315
   End
   Begin VB.TextBox txtDay 
      Height          =   435
      Index           =   0
      Left            =   4470
      TabIndex        =   46
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Label3 
      Caption         =   "询价单编号"
      Height          =   225
      Left            =   12360
      TabIndex        =   95
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   4
      Left            =   4560
      TabIndex        =   94
      Top             =   8700
      Width           =   945
   End
   Begin VB.Label lblQM 
      Caption         =   "lblQm"
      Height          =   225
      Index           =   4
      Left            =   4650
      TabIndex        =   93
      Top             =   7980
      Width           =   915
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   3
      Left            =   3480
      TabIndex        =   91
      Top             =   8700
      Width           =   945
   End
   Begin VB.Label lblQM 
      Caption         =   "lblQm"
      Height          =   225
      Index           =   3
      Left            =   3570
      TabIndex        =   90
      Top             =   7980
      Width           =   915
   End
   Begin VB.Label lblQM 
      Caption         =   "lblQm"
      Height          =   225
      Index           =   2
      Left            =   2490
      TabIndex        =   88
      Top             =   7980
      Width           =   915
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   2
      Left            =   2400
      TabIndex        =   87
      Top             =   8700
      Width           =   945
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   1320
      TabIndex        =   75
      Top             =   8700
      Width           =   945
   End
   Begin VB.Label lblQM 
      Caption         =   "lblQm"
      Height          =   225
      Index           =   1
      Left            =   1410
      TabIndex        =   74
      Top             =   7980
      Width           =   915
   End
   Begin VB.Label lblQM 
      Caption         =   "lblQm"
      Height          =   225
      Index           =   0
      Left            =   330
      TabIndex        =   71
      Top             =   7980
      Width           =   915
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   240
      TabIndex        =   70
      Top             =   8700
      Width           =   945
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   14
      X1              =   4470
      X2              =   4695
      Y1              =   7110
      Y2              =   7110
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   13
      X1              =   4470
      X2              =   4695
      Y1              =   6690
      Y2              =   6690
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   12
      X1              =   4470
      X2              =   4695
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   11
      X1              =   4470
      X2              =   4695
      Y1              =   5790
      Y2              =   5790
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   10
      X1              =   4470
      X2              =   4695
      Y1              =   5340
      Y2              =   5340
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   9
      X1              =   4470
      X2              =   4695
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   8
      X1              =   4470
      X2              =   4695
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   7
      X1              =   4470
      X2              =   4695
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   6
      X1              =   4470
      X2              =   4695
      Y1              =   3570
      Y2              =   3570
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   5
      X1              =   4470
      X2              =   4695
      Y1              =   3090
      Y2              =   3090
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   4
      X1              =   4470
      X2              =   4695
      Y1              =   2670
      Y2              =   2670
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   3
      X1              =   4470
      X2              =   4695
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   2
      X1              =   4470
      X2              =   4695
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   1
      X1              =   4470
      X2              =   4695
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line liD 
      BorderWidth     =   3
      Index           =   15
      X1              =   4470
      X2              =   4695
      Y1              =   7530
      Y2              =   7530
   End
   Begin VB.Line Line4 
      Index           =   20
      X1              =   13635
      X2              =   13635
      Y1              =   600
      Y2              =   7770
   End
   Begin VB.Line Line4 
      Index           =   19
      X1              =   13161
      X2              =   13161
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   18
      X1              =   12702
      X2              =   12702
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   17
      X1              =   12243
      X2              =   12243
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   16
      X1              =   11784
      X2              =   11784
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   15
      X1              =   11325
      X2              =   11325
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   14
      X1              =   10866
      X2              =   10866
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   13
      X1              =   10407
      X2              =   10407
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   12
      X1              =   9948
      X2              =   9948
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   11
      X1              =   9489
      X2              =   9489
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   10
      X1              =   9000
      X2              =   9000
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   8
      X1              =   8112
      X2              =   8112
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   7
      X1              =   7653
      X2              =   7653
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   6
      X1              =   7200
      X2              =   7200
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   5
      X1              =   6735
      X2              =   6735
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   4
      X1              =   6276
      X2              =   6276
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   3
      X1              =   5817
      X2              =   5817
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   5358
      X2              =   5358
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   4899
      X2              =   4899
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line4 
      Index           =   0
      X1              =   4440
      X2              =   4440
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   7215
      Left            =   240
      Top             =   570
      Width           =   14055
   End
   Begin VB.Line Line4 
      Index           =   9
      X1              =   8571
      X2              =   8571
      Y1              =   570
      Y2              =   7740
   End
   Begin VB.Line Line3 
      X1              =   270
      X2              =   14280
      Y1              =   7770
      Y2              =   7770
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "15"
      Height          =   375
      Index           =   14
      Left            =   270
      TabIndex        =   45
      Top             =   7350
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "14"
      Height          =   375
      Index           =   13
      Left            =   270
      TabIndex        =   44
      Top             =   6895
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "13"
      Height          =   375
      Index           =   12
      Left            =   270
      TabIndex        =   43
      Top             =   6450
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "12"
      Height          =   375
      Index           =   11
      Left            =   270
      TabIndex        =   42
      Top             =   6005
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "11"
      Height          =   375
      Index           =   10
      Left            =   270
      TabIndex        =   41
      Top             =   5560
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "10"
      Height          =   375
      Index           =   9
      Left            =   270
      TabIndex        =   40
      Top             =   5115
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "9"
      Height          =   375
      Index           =   8
      Left            =   270
      TabIndex        =   39
      Top             =   4670
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "8"
      Height          =   375
      Index           =   7
      Left            =   270
      TabIndex        =   38
      Top             =   4225
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "7"
      Height          =   375
      Index           =   6
      Left            =   270
      TabIndex        =   37
      Top             =   3780
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "6"
      Height          =   375
      Index           =   5
      Left            =   270
      TabIndex        =   36
      Top             =   3335
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "5"
      Height          =   375
      Index           =   4
      Left            =   270
      TabIndex        =   35
      Top             =   2890
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "4"
      Height          =   375
      Index           =   3
      Left            =   270
      TabIndex        =   34
      Top             =   2445
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "3"
      Height          =   375
      Index           =   2
      Left            =   270
      TabIndex        =   33
      Top             =   2000
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "2"
      Height          =   375
      Index           =   1
      Left            =   270
      TabIndex        =   32
      Top             =   1555
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   375
      Index           =   0
      Left            =   270
      TabIndex        =   31
      Top             =   1110
      Width           =   405
   End
   Begin VB.Label Label1 
      Caption         =   "日期"
      Height          =   285
      Left            =   300
      TabIndex        =   29
      Top             =   690
      Width           =   615
   End
   Begin VB.Line Line2 
      X1              =   210
      X2              =   14250
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   240
      X2              =   14280
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   240
      X2              =   14280
      Y1              =   6855
      Y2              =   6855
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   240
      X2              =   14280
      Y1              =   6420
      Y2              =   6420
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   240
      X2              =   14280
      Y1              =   5970
      Y2              =   5970
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   240
      X2              =   14295
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   210
      X2              =   14250
      Y1              =   5070
      Y2              =   5070
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   210
      X2              =   14250
      Y1              =   4620
      Y2              =   4620
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   210
      X2              =   14250
      Y1              =   4185
      Y2              =   4185
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   210
      X2              =   14250
      Y1              =   3735
      Y2              =   3735
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   210
      X2              =   14265
      Y1              =   3285
      Y2              =   3285
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   210
      X2              =   14250
      Y1              =   2835
      Y2              =   2835
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   210
      X2              =   14250
      Y1              =   2385
      Y2              =   2385
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   210
      X2              =   14250
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   210
      X2              =   14250
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   210
      X2              =   14280
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Label lblY 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   7
      Left            =   14580
      TabIndex        =   13
      Top             =   5160
      Width           =   315
   End
   Begin VB.Label lblY 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   14580
      TabIndex        =   12
      Top             =   4590
      Width           =   315
   End
   Begin VB.Label lblY 
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   14580
      TabIndex        =   11
      Top             =   4050
      Width           =   315
   End
   Begin VB.Label lblY 
      BackColor       =   &H00FF00FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   14580
      TabIndex        =   10
      Top             =   3510
      Width           =   315
   End
   Begin VB.Label lblY 
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   14580
      TabIndex        =   9
      Top             =   2910
      Width           =   315
   End
   Begin VB.Label lblY 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   14580
      TabIndex        =   8
      Top             =   2340
      Width           =   315
   End
   Begin VB.Label lblY 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   14580
      TabIndex        =   7
      Top             =   1800
      Width           =   315
   End
   Begin VB.Label lblY 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   0
      Left            =   14580
      TabIndex        =   6
      Top             =   1260
      Width           =   315
   End
   Begin VB.Label Label8 
      Caption         =   "组长"
      Height          =   225
      Left            =   9840
      TabIndex        =   5
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label7 
      Caption         =   "项目名称"
      Height          =   255
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   735
   End
   Begin VB.Label Label25 
      Caption         =   "合同编号"
      Height          =   225
      Left            =   5220
      TabIndex        =   2
      Top             =   90
      Width           =   855
   End
End
Attribute VB_Name = "frmGJW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim LIT As Integer '目前选择的线条序号
Dim MX As Long
Dim MY As Long



Private Sub cmdBack_Click()
Me.Visible = False
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf frmGjwV.Visible = True Then
    frmGjwV.Enabled = True
    frmGjwV.ZOrder 0
End If

End Sub



Private Sub cmdMod_Click()
If lblLc.Caption = 2 Or lblPwf.Caption = "True" Then '工程组长签字后,将不得再修改
    Exit Sub
End If
If Not (txtZu.Text = mod1.DName And lblUid.Caption = mod1.DHid) Then
    Exit Sub
End If

cmdSave.Enabled = True
End Sub

Private Sub cmdQm_Click(Index As Integer)
Dim oo As Integer
Dim tt As String
Dim Zid As Long
On Error Resume Next
If cmdQm(Index).Caption <> "" Or lblLcRen.Caption = "" Then
    Exit Sub
End If
If Not (lblLcRen.Caption = mod1.DName And lblLcUid.Caption = mod1.DHid) And lblLc.Caption <> (Index + 1) Then
    Exit Sub
End If

If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If


If lblLcUid.Caption <> mod1.DHid Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If

Dim Zi As Integer
Zi = MsgBox("是否确认签字?", vbYesNo)
If Zi = vbNo Then Exit Sub

    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "xtzxAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@yid").Value = 73
    mod1.cmd.Parameters("@lc").Value = lblLc.Caption
    mod1.cmd.Parameters("@bh").Value = lblGid.Caption
    mod1.cmd.Parameters("@ywy").Value = mod1.DName
    mod1.cmd.Parameters("@uid").Value = mod1.DHid
    mod1.cmd.Parameters("@BZ").Value = ""
    mod1.cmd.Execute
    Zid = mod1.cmd.Parameters("@Zid").Value
    Set cmd = Nothing

cmdQm(Index).Caption = mod1.DName
lblTm(Index).Caption = mod1.DQda
lblLc.Caption = lblLc.Caption + 1
lblLcRen.Caption = ""
lblLcUid.Caption = ""
If Dialog.Visible = True Then
    Call mod1.refEnvent(1)

End If
End Sub

Private Sub cmdSave_Click()
Dim oo As Integer
Dim tt As String
On Error Resume Next
tt = "select * from gjw where gid=" & Val(lblGid.Caption)
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
mod1.HTP.Update "Gdate", txtDate.Text
For oo = 0 To 20
    If txtDay(oo).ToolTipText <> "" Then
        mod1.HTP.Update "Gday" & oo, txtDay(oo).ToolTipText
    Else
        mod1.HTP.Update "Gday" & oo, Null
    End If
Next
'记录工作内容与线条
For oo = 1 To 15
    mod1.HTP.Update "nr" & oo, txtNr(oo).Text
    mod1.HTP.Update "x1" & oo, liD(oo).X1
    mod1.HTP.Update "x2" & oo, liD(oo).x2
    mod1.HTP.Update "lcolor" & oo, liD(oo).BorderColor
Next
mod1.HTP.UpdateBatch
cmdSave.Enabled = False
cmdMod.Enabled = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    liD(Index).Visible = False
End If
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight





End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oo As Integer
If Y < 570 Or Y > 7770 Or X < 4410 Or X > 14250 Then
    Exit Sub
End If
If Y > txtNr(15).Top Then
    LIT = 15
Else
    For oo = 14 To 1 Step -1
        If Y > txtNr(oo).Top And Y < txtNr(oo + 1).Top Then
            LIT = oo
            Exit For
        End If
    Next
End If
    
If Shift = 2 And Button = 1 Then
    liD(LIT).Visible = True
    liD(LIT).X1 = X
    liD(LIT).x2 = X + 10
ElseIf Shift = 2 And Button = 2 And liD(LIT).Visible = True Then
    liD(LIT).x2 = X
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MX = X
MY = Y
End Sub

Private Sub lblY_DblClick(Index As Integer)
    liD(LIT).BorderColor = lblY(Index).BackColor
End Sub


Private Sub txtDay_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    On Error GoTo dayERRor
    
    txtDay(Index).ToolTipText = DateSerial(Year(txtDay(Index).Text), Month(txtDay(Index).Text), Day(txtDay(Index).Text))
    txtDay(Index).Text = Day(txtDay(Index).Text)

        txtDay(Index).Text = txtDay(Index).Text & "日"

End If
Exit Sub
dayERRor:
txtDay(Index).Text = ""
txtDay(Index).ToolTipText = ""
End Sub


Private Sub txtDay_LostFocus(Index As Integer)
If KeyCode = 13 Then
    On Error GoTo dayERRor1
    
    txtDay(Index).ToolTipText = DateSerial(Year(txtDay(Index).Text), Month(txtDay(Index).Text), Day(txtDay(Index).Text))
    txtDay(Index).Text = Day(txtDay(Index).Text)

        txtDay(Index).Text = txtDay(Index).Text & "日"

End If
Exit Sub
dayERRor1:
txtDay(Index).Text = ""
txtDay(Index).ToolTipText = ""
End Sub


Private Sub txtNr_Click(Index As Integer)
'LIT = Index
'MsgBox MY
End Sub


