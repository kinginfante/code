VERSION 5.00
Begin VB.Form frmWBXT1 
   Caption         =   "����ѯ��ϵͳ"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14220
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   14220
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   630
      Top             =   570
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   570
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   12330
      Picture         =   "frmWBXT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   100
      ToolTipText     =   "�޸�"
      Top             =   7950
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   12840
      Picture         =   "frmWBXT.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   99
      ToolTipText     =   "����"
      Top             =   7950
      Width           =   465
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   13740
      Picture         =   "frmWBXT.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   98
      ToolTipText     =   "����"
      Top             =   7920
      Width           =   465
   End
   Begin VB.CommandButton cmdD 
      Enabled         =   0   'False
      Height          =   405
      Left            =   13230
      Picture         =   "frmWBXT.frx":0A76
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   7920
      Width           =   465
   End
   Begin VB.Frame Frame1 
      Caption         =   "�廯﮻���ά��"
      Height          =   8475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14235
      Begin VB.TextBox J17 
         Height          =   270
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   113
         Text            =   "Text5"
         Top             =   6795
         Width           =   1065
      End
      Begin VB.TextBox H17 
         Height          =   270
         Left            =   10650
         Locked          =   -1  'True
         TabIndex        =   112
         Text            =   "Text5"
         Top             =   6795
         Width           =   1065
      End
      Begin VB.TextBox F17 
         Height          =   270
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   111
         Text            =   "Text5"
         Top             =   6795
         Width           =   1065
      End
      Begin VB.ComboBox C8 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0C00
         Left            =   5010
         List            =   "frmWBXT.frx":0C0D
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   2775
         Width           =   1005
      End
      Begin VB.TextBox D7 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   109
         Text            =   "Text5"
         Top             =   2430
         Width           =   1065
      End
      Begin VB.ComboBox C9 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0C29
         Left            =   5010
         List            =   "frmWBXT.frx":0C36
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   3145
         Width           =   1005
      End
      Begin VB.TextBox D8 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   107
         Text            =   "Text5"
         Top             =   2775
         Width           =   1065
      End
      Begin VB.TextBox D6 
         Height          =   300
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   106
         Top             =   2040
         Width           =   1005
      End
      Begin VB.ComboBox I3 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0C52
         Left            =   11880
         List            =   "frmWBXT.frx":0C89
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox G3 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0CC7
         Left            =   9570
         List            =   "frmWBXT.frx":0CFE
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox E3 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0D3C
         Left            =   7290
         List            =   "frmWBXT.frx":0D73
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox C3 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0DB1
         Left            =   5010
         List            =   "frmWBXT.frx":0DE8
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   960
         Width           =   1005
      End
      Begin VB.CommandButton cmdJi 
         Caption         =   "����"
         Height          =   345
         Left            =   11790
         TabIndex        =   101
         Top             =   7950
         Width           =   525
      End
      Begin VB.TextBox J10 
         Height          =   270
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   96
         Text            =   "Text5"
         Top             =   3513
         Width           =   1065
      End
      Begin VB.TextBox H10 
         Height          =   270
         Left            =   10650
         Locked          =   -1  'True
         TabIndex        =   95
         Text            =   "Text5"
         Top             =   3513
         Width           =   1065
      End
      Begin VB.ComboBox I10 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0E26
         Left            =   11880
         List            =   "frmWBXT.frx":0E33
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   3513
         Width           =   1005
      End
      Begin VB.TextBox D19 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   91
         Text            =   "Text5"
         Top             =   7500
         Width           =   1065
      End
      Begin VB.TextBox D18 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   90
         Text            =   "Text5"
         Top             =   7155
         Width           =   1065
      End
      Begin VB.TextBox J8 
         Height          =   270
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   89
         Text            =   "Text5"
         Top             =   2775
         Width           =   1065
      End
      Begin VB.TextBox J14 
         Height          =   270
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   88
         Text            =   "Text5"
         Top             =   4985
         Width           =   1065
      End
      Begin VB.TextBox J13 
         Height          =   270
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   87
         Text            =   "Text5"
         Top             =   4617
         Width           =   1065
      End
      Begin VB.TextBox J12 
         Height          =   270
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   86
         Text            =   "Text5"
         Top             =   4249
         Width           =   1065
      End
      Begin VB.TextBox J11 
         Height          =   270
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   85
         Text            =   "Text5"
         Top             =   3881
         Width           =   1065
      End
      Begin VB.TextBox J9 
         Height          =   270
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   84
         Text            =   "Text6"
         Top             =   3145
         Width           =   1065
      End
      Begin VB.TextBox H8 
         Height          =   270
         Left            =   10650
         Locked          =   -1  'True
         TabIndex        =   83
         Text            =   "Text5"
         Top             =   2775
         Width           =   1065
      End
      Begin VB.TextBox H14 
         Height          =   270
         Left            =   10650
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "Text5"
         Top             =   4985
         Width           =   1065
      End
      Begin VB.TextBox H13 
         Height          =   270
         Left            =   10650
         Locked          =   -1  'True
         TabIndex        =   81
         Text            =   "Text5"
         Top             =   4617
         Width           =   1065
      End
      Begin VB.TextBox H12 
         Height          =   270
         Left            =   10650
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "Text5"
         Top             =   4249
         Width           =   1065
      End
      Begin VB.TextBox H11 
         Height          =   270
         Left            =   10650
         Locked          =   -1  'True
         TabIndex        =   79
         Text            =   "Text5"
         Top             =   3881
         Width           =   1065
      End
      Begin VB.TextBox H9 
         Height          =   270
         Left            =   10650
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "Text6"
         Top             =   3145
         Width           =   1065
      End
      Begin VB.TextBox F8 
         Height          =   270
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "Text5"
         Top             =   2775
         Width           =   1065
      End
      Begin VB.TextBox F14 
         Height          =   270
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Text5"
         Top             =   4985
         Width           =   1065
      End
      Begin VB.TextBox F13 
         Height          =   270
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Text5"
         Top             =   4617
         Width           =   1065
      End
      Begin VB.TextBox F12 
         Height          =   270
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   74
         Text            =   "Text5"
         Top             =   4249
         Width           =   1065
      End
      Begin VB.TextBox F11 
         Height          =   270
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "Text5"
         Top             =   3881
         Width           =   1065
      End
      Begin VB.TextBox F10 
         Height          =   270
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text5"
         Top             =   3513
         Width           =   1065
      End
      Begin VB.TextBox F9 
         Height          =   270
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text6"
         Top             =   3145
         Width           =   1065
      End
      Begin VB.TextBox D17 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   6795
         Width           =   1065
      End
      Begin VB.TextBox D16 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "Text5"
         Top             =   6450
         Width           =   1065
      End
      Begin VB.TextBox D15 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "Text5"
         Top             =   6105
         Width           =   1065
      End
      Begin VB.TextBox D14 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "Text5"
         Top             =   4995
         Width           =   1065
      End
      Begin VB.TextBox D13 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "Text5"
         Top             =   4625
         Width           =   1065
      End
      Begin VB.TextBox D12 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text5"
         Top             =   4255
         Width           =   1065
      End
      Begin VB.TextBox D11 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Text5"
         Top             =   3885
         Width           =   1065
      End
      Begin VB.TextBox D10 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "Text5"
         Top             =   3515
         Width           =   1065
      End
      Begin VB.TextBox D9 
         Height          =   270
         Left            =   6090
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text6"
         Top             =   3145
         Width           =   1065
      End
      Begin VB.ComboBox G9 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0E4F
         Left            =   9585
         List            =   "frmWBXT.frx":0E5C
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   3145
         Width           =   1005
      End
      Begin VB.ComboBox E9 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0E78
         Left            =   7305
         List            =   "frmWBXT.frx":0E85
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   3145
         Width           =   1005
      End
      Begin VB.ComboBox I8 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0EA1
         Left            =   11880
         List            =   "frmWBXT.frx":0EAB
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   2775
         Width           =   1005
      End
      Begin VB.ComboBox I14 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0EBD
         Left            =   11880
         List            =   "frmWBXT.frx":0EC7
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   4985
         Width           =   1005
      End
      Begin VB.ComboBox I13 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0ED9
         Left            =   11880
         List            =   "frmWBXT.frx":0EE3
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   4617
         Width           =   1005
      End
      Begin VB.ComboBox I12 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0EF5
         Left            =   11880
         List            =   "frmWBXT.frx":0EFF
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   4249
         Width           =   1005
      End
      Begin VB.ComboBox I11 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0F11
         Left            =   11880
         List            =   "frmWBXT.frx":0F1B
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   3881
         Width           =   1005
      End
      Begin VB.ComboBox I9 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0F2D
         Left            =   11880
         List            =   "frmWBXT.frx":0F3A
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   3145
         Width           =   1005
      End
      Begin VB.ComboBox G8 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0F56
         Left            =   9585
         List            =   "frmWBXT.frx":0F60
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   2775
         Width           =   1005
      End
      Begin VB.ComboBox G14 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0F72
         Left            =   9585
         List            =   "frmWBXT.frx":0F7C
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   4985
         Width           =   1005
      End
      Begin VB.ComboBox G13 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0F8E
         Left            =   9585
         List            =   "frmWBXT.frx":0F98
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   4617
         Width           =   1005
      End
      Begin VB.ComboBox G12 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0FAA
         Left            =   9585
         List            =   "frmWBXT.frx":0FB4
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   4249
         Width           =   1005
      End
      Begin VB.ComboBox G11 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0FC6
         Left            =   9585
         List            =   "frmWBXT.frx":0FD0
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   3881
         Width           =   1005
      End
      Begin VB.ComboBox G10 
         Height          =   300
         ItemData        =   "frmWBXT.frx":0FE2
         Left            =   9585
         List            =   "frmWBXT.frx":0FEF
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   3513
         Width           =   1005
      End
      Begin VB.ComboBox E8 
         Height          =   300
         ItemData        =   "frmWBXT.frx":100B
         Left            =   7305
         List            =   "frmWBXT.frx":1015
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   2775
         Width           =   1005
      End
      Begin VB.ComboBox E14 
         Height          =   300
         ItemData        =   "frmWBXT.frx":1027
         Left            =   7305
         List            =   "frmWBXT.frx":1031
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   4995
         Width           =   1005
      End
      Begin VB.ComboBox E13 
         Height          =   300
         ItemData        =   "frmWBXT.frx":1043
         Left            =   7305
         List            =   "frmWBXT.frx":104D
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   4625
         Width           =   1005
      End
      Begin VB.ComboBox E12 
         Height          =   300
         ItemData        =   "frmWBXT.frx":105F
         Left            =   7305
         List            =   "frmWBXT.frx":1069
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   4255
         Width           =   1005
      End
      Begin VB.ComboBox E11 
         Height          =   300
         ItemData        =   "frmWBXT.frx":107B
         Left            =   7305
         List            =   "frmWBXT.frx":1085
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   3885
         Width           =   1005
      End
      Begin VB.ComboBox E10 
         Height          =   300
         ItemData        =   "frmWBXT.frx":1097
         Left            =   7305
         List            =   "frmWBXT.frx":10A4
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   3515
         Width           =   1005
      End
      Begin VB.ComboBox C16 
         Height          =   300
         ItemData        =   "frmWBXT.frx":10C0
         Left            =   5010
         List            =   "frmWBXT.frx":10CA
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   6450
         Width           =   1005
      End
      Begin VB.ComboBox C14 
         Height          =   300
         ItemData        =   "frmWBXT.frx":10DC
         Left            =   5010
         List            =   "frmWBXT.frx":10E6
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   4995
         Width           =   1005
      End
      Begin VB.ComboBox C13 
         Height          =   300
         ItemData        =   "frmWBXT.frx":10F8
         Left            =   5010
         List            =   "frmWBXT.frx":1102
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   4625
         Width           =   1005
      End
      Begin VB.ComboBox C12 
         Height          =   300
         ItemData        =   "frmWBXT.frx":1114
         Left            =   5010
         List            =   "frmWBXT.frx":111E
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   4255
         Width           =   1005
      End
      Begin VB.ComboBox C11 
         Height          =   300
         ItemData        =   "frmWBXT.frx":1130
         Left            =   5010
         List            =   "frmWBXT.frx":113A
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   3885
         Width           =   1005
      End
      Begin VB.ComboBox C10 
         Height          =   300
         ItemData        =   "frmWBXT.frx":114C
         Left            =   5010
         List            =   "frmWBXT.frx":1159
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   3515
         Width           =   1005
      End
      Begin VB.ComboBox C6 
         Height          =   300
         ItemData        =   "frmWBXT.frx":1175
         Left            =   5010
         List            =   "frmWBXT.frx":11A9
         TabIndex        =   32
         Text            =   "C6"
         Top             =   2040
         Width           =   1005
      End
      Begin VB.TextBox J6 
         Height          =   300
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2040
         Width           =   1005
      End
      Begin VB.TextBox H6 
         Height          =   300
         Left            =   10635
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2040
         Width           =   1005
      End
      Begin VB.TextBox F6 
         Height          =   300
         Left            =   8385
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2040
         Width           =   1005
      End
      Begin VB.ComboBox I6 
         Height          =   300
         ItemData        =   "frmWBXT.frx":1203
         Left            =   11880
         List            =   "frmWBXT.frx":1237
         TabIndex        =   28
         Text            =   "I6"
         Top             =   2040
         Width           =   1005
      End
      Begin VB.ComboBox G6 
         Height          =   300
         ItemData        =   "frmWBXT.frx":1291
         Left            =   9585
         List            =   "frmWBXT.frx":12C5
         TabIndex        =   27
         Text            =   "G6"
         Top             =   2040
         Width           =   1005
      End
      Begin VB.ComboBox E6 
         Height          =   300
         ItemData        =   "frmWBXT.frx":131F
         Left            =   7305
         List            =   "frmWBXT.frx":1353
         TabIndex        =   26
         Text            =   "E6"
         Top             =   2040
         Width           =   1005
      End
      Begin VB.ComboBox I5 
         Height          =   300
         ItemData        =   "frmWBXT.frx":13AD
         Left            =   11880
         List            =   "frmWBXT.frx":13C9
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1680
         Width           =   1005
      End
      Begin VB.ComboBox G5 
         Height          =   300
         ItemData        =   "frmWBXT.frx":13E5
         Left            =   9585
         List            =   "frmWBXT.frx":1401
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1680
         Width           =   1005
      End
      Begin VB.ComboBox E5 
         Height          =   300
         ItemData        =   "frmWBXT.frx":141D
         Left            =   7305
         List            =   "frmWBXT.frx":1439
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1680
         Width           =   1005
      End
      Begin VB.ComboBox C5 
         Height          =   300
         ItemData        =   "frmWBXT.frx":1455
         Left            =   5010
         List            =   "frmWBXT.frx":1471
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1680
         Width           =   1005
      End
      Begin VB.ComboBox I4 
         Height          =   300
         ItemData        =   "frmWBXT.frx":148D
         Left            =   11880
         List            =   "frmWBXT.frx":14C4
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1320
         Width           =   1005
      End
      Begin VB.ComboBox G4 
         Height          =   300
         ItemData        =   "frmWBXT.frx":1502
         Left            =   9585
         List            =   "frmWBXT.frx":1539
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1320
         Width           =   1005
      End
      Begin VB.ComboBox E4 
         Height          =   300
         ItemData        =   "frmWBXT.frx":1577
         Left            =   7305
         List            =   "frmWBXT.frx":15AE
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1320
         Width           =   1005
      End
      Begin VB.ComboBox C4 
         Height          =   300
         ItemData        =   "frmWBXT.frx":15EC
         Left            =   5010
         List            =   "frmWBXT.frx":1623
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   14220
         Y1              =   7830
         Y2              =   7830
      End
      Begin VB.Line Line6 
         X1              =   0
         X2              =   14220
         Y1              =   6030
         Y2              =   6030
      End
      Begin VB.Label Label25 
         Caption         =   $"frmWBXT.frx":1661
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2010
         TabIndex        =   93
         Top             =   7560
         Width           =   2055
      End
      Begin VB.Label Label24 
         Caption         =   $"frmWBXT.frx":1678
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2010
         TabIndex        =   92
         Top             =   7140
         Width           =   2115
      End
      Begin VB.Line Line5 
         X1              =   14220
         X2              =   0
         Y1              =   7470
         Y2              =   7470
      End
      Begin VB.Line Line4 
         X1              =   14220
         X2              =   0
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Line Line3 
         X1              =   7230
         X2              =   0
         Y1              =   6750
         Y2              =   6750
      End
      Begin VB.Line Line2 
         X1              =   7230
         X2              =   0
         Y1              =   6390
         Y2              =   6390
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   960
         TabIndex        =   61
         Top             =   6795
         Width           =   3435
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Ѳ��"
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   960
         TabIndex        =   60
         Top             =   6465
         Width           =   3435
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "����"
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   960
         TabIndex        =   59
         Top             =   6180
         Width           =   3435
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   11820
         X2              =   11820
         Y1              =   90
         Y2              =   7080
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   9510
         X2              =   9510
         Y1              =   90
         Y2              =   7080
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   4680
         X2              =   4680
         Y1              =   90
         Y2              =   7860
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   7260
         X2              =   7260
         Y1              =   90
         Y2              =   7830
      End
      Begin VB.Label Label20 
         Caption         =   "ѡ��              ���    ѡ��            ���     ѡ��          ���        ѡ��         ��� "
         Height          =   255
         Left            =   5070
         TabIndex        =   17
         Top             =   630
         Width           =   8895
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmWBXT.frx":168F
         Height          =   285
         Left            =   2130
         TabIndex        =   16
         Top             =   5025
         Width           =   2265
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmWBXT.frx":1699
         Height          =   285
         Left            =   2130
         TabIndex        =   15
         Top             =   4650
         Width           =   2265
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "�������Ʊ�������"
         Height          =   285
         Left            =   2130
         TabIndex        =   14
         Top             =   4275
         Width           =   2265
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "ȼ��¯����"
         Height          =   285
         Left            =   2130
         TabIndex        =   13
         Top             =   3915
         Width           =   2265
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "��Һ����"
         Height          =   285
         Left            =   2130
         TabIndex        =   12
         Top             =   3566
         Width           =   2265
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmWBXT.frx":16A7
         Height          =   285
         Left            =   2130
         TabIndex        =   11
         Top             =   3198
         Width           =   2265
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "��������е��ϴ"
         Height          =   285
         Left            =   2130
         TabIndex        =   10
         Top             =   2830
         Width           =   2265
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmWBXT.frx":16BB
         Height          =   285
         Left            =   2130
         TabIndex        =   9
         Top             =   2462
         Width           =   2265
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "����ʽֱȼʽ"
         Height          =   285
         Left            =   2130
         TabIndex        =   8
         Top             =   2094
         Width           =   2265
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "������Сkw"
         Height          =   285
         Left            =   2130
         TabIndex        =   7
         Top             =   1726
         Width           =   2265
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "ÿ̨����õ�����"
         Height          =   285
         Left            =   2130
         TabIndex        =   6
         Top             =   1358
         Width           =   2265
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "��������"
         Height          =   285
         Left            =   2130
         TabIndex        =   5
         Top             =   990
         Width           =   2265
      End
      Begin VB.Label Label4 
         Caption         =   "��ȱ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   210
         TabIndex        =   4
         Top             =   3420
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "ѯ�۲���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   300
         TabIndex        =   3
         Top             =   1050
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   "����"
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
         Left            =   2520
         TabIndex        =   2
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmWBXT1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Qing()
C3.Text = 0: E3.Text = 0: G3.Text = 0: I3.Text = 0
C4.Text = 1: E4.Text = 1: G4.Text = 1: I4.Text = 1
C5.Text = 1: E5.Text = 1: G5.Text = 1: I5.Text = 1
C7.Text = "": E7.Text = "": G7.Text = "": I7.Text = ""
D8.Text = 500
D9.Text = "": F9.Text = "": H9.Text = "": J9.Text = ""
D10.Text = "": F10.Text = "": H10.Text = "": J10.Text = ""
D11.Text = "": F11.Text = "": H11.Text = "": J11.Text = ""
D12.Text = "": F12.Text = "": H12.Text = "": J12.Text = ""
D13.Text = "": F13.Text = "": H13.Text = "": J13.Text = ""
D14.Text = "": F14.Text = "": H14.Text = "": J14.Text = ""
D15.Text = "": F15.Text = "": H15.Text = "": J15.Text = ""
D16.Text = "": F16.Text = "": H16.Text = "": J16.Text = ""
D17.Text = "": F17.Text = "": H17.Text = "": J17.Text = ""
D18.Text = "": D19.Text = "": D20.Text = "": D21.Text = "": D22.Text = ""
C9.Text = "����Ҫ": E9.Text = "����Ҫ": G9.Text = "����Ҫ": I9.Text = "����Ҫ"
C10.Text = "����Ҫ": E10.Text = "����Ҫ": G10.Text = "����Ҫ": I10.Text = "����Ҫ"
C11.Text = "����Ҫ": E11.Text = "����Ҫ": G11.Text = "����Ҫ": I11.Text = "����Ҫ"
C12.Text = "����Ҫ": E12.Text = "����Ҫ": G12.Text = "����Ҫ": I12.Text = "����Ҫ"
C13.Text = "����Ҫ": E13.Text = "����Ҫ": G13.Text = "����Ҫ": I13.Text = "����Ҫ"
C14.Text = "����Ҫ": E14.Text = "����Ҫ": G14.Text = "����Ҫ": I14.Text = "����Ҫ"
C15.Text = "����Ҫ": E15.Text = "����Ҫ": G15.Text = "����Ҫ": I15.Text = "����Ҫ"
C16.Text = "����Ҫ": E16.Text = "����Ҫ": G16.Text = "����Ҫ": I16.Text = "����Ҫ"
C17.Text = "����Ҫ": E17.Text = "����Ҫ": G17.Text = "����Ҫ": I17.Text = "����Ҫ"
C19.Text = 5
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
End Sub



Private Sub cmdJi_Click()
If Val(C4.Text) = 0 And Val(E4.Text) = 0 And Val(G4.Text) = 0 And Val(I4.Text) = 0 Then
    MsgBox "��ȷ����������!"
End If
C7 = C6 * C5
E7 = E6 * E5
G7 = G6 * G5
I7 = I6 * I5
If C9 = "��һ��" Then
    If C2 = "ˮ��" Then
        D9 = Round(300 * C4 * (1 + C7 / 1050), 0)
    Else
        D9 = 0
    End If
ElseIf C9 = "������" Then
    If C2 = "ˮ��" Then
        D9 = Round(300 * C4 * (1 + C7 / 1050) * 1.1, 0)
    Else
        D9 = 0
    End If
Else
    D9 = 0
End If

If E9 = "��һ��" Then
    If E2 = "ˮ��" Then
        F9 = Round(300 * E4 * (1 + E7 / 1050), 0)
    Else
        F9 = 0
    End If
ElseIf E9 = "������" Then
    If E2 = "ˮ��" Then
        F9 = Round(300 * E4 * (1 + E7 / 1050) * 1.1, 0)
    Else
        F9 = 0
    End If
Else
    F9 = 0
End If

If G9 = "��һ��" Then
    If G2 = "ˮ��" Then
        H9 = Round(300 * G4 * (1 + G7 / 1050), 0)
    Else
        H9 = 0
    End If
ElseIf G9 = "������" Then
    If G2 = "ˮ��" Then
        H9 = Round(300 * G4 * (1 + G7 / 1050) * 1.1, 0)
    Else
        H9 = 0
    End If
Else
    H9 = 0
End If

If I9 = "��һ��" Then
    If I2 = "ˮ��" Then
        J9 = Round(300 * I4 * (1 + I7 / 1050), 0)
    Else
        J9 = 0
    End If
ElseIf I9 = "������" Then
    If I2 = "ˮ��" Then
        J9 = Round(300 * I4 * (1 + I7 / 1050) * 1.1, 0)
    Else
        J9 = 0
    End If
Else
    J9 = 0
End If

If C10 = "��һ��" Then
    If C2 = "ˮ��" Then
        D10 = Round(300 * C4 * (1 + C7 / 1050), 0)
    Else
        D10 = 0
    End If
ElseIf C10 = "������" Then
    If C2 = "ˮ��" Then
        D10 = Round(300 * C4 * (1 + C7 / 1050) * 1.1, 0)
    Else
        D10 = 0
    End If
Else
    D10 = 0
End If

If E10 = "��һ��" Then
    If E2 = "ˮ��" Then
        F10 = Round(300 * E4 * (1 + E7 / 1050), 0)
    Else
        F10 = 0
    End If
ElseIf E10 = "������" Then
    If E2 = "ˮ��" Then
        F10 = Round(300 * E4 * (1 + E7 / 1050) * 1.1, 0)
    Else
        F10 = 0
    End If
Else
    F10 = 0
End If
If G10 = "��һ��" Then
    If G2 = "ˮ��" Then
        H10 = Round(300 * G4 * (1 + G7 / 1050), 0)
    Else
        H10 = 0
    End If
ElseIf G10 = "������" Then
    If G2 = "ˮ��" Then
        H10 = Round(300 * G4 * (1 + G7 / 1050) * 1.1, 0)
    Else
        H10 = 0
    End If
Else
    H10 = 0
End If
If I10 = "��һ��" Then
    If I2 = "ˮ��" Then
        J10 = Round(300 * I4 * (1 + I7 / 1050), 0)
    Else
        J10 = 0
    End If
ElseIf I10 = "������" Then
    If I2 = "ˮ��" Then
        J10 = Round(300 * I4 * (1 + I7 / 1050) * 1.1, 0)
    Else
        J10 = 0
    End If
Else
    J10 = 0
End If
If C2 = "ˮ��" Then
    D11 = 0
Else
    If C11 = "��Ҫ" Then
        D11 = Round(600 * C4 * C7 / 350, 0)
    Else
        D11 = 0
    End If
End If
If E2 = "ˮ��" Then
    F11 = 0
Else
    If E11 = "��Ҫ" Then
        F11 = Round(600 * E4 * E7 / 350, 0)
    Else
        F11 = 0
    End If
End If
If G2 = "ˮ��" Then
    H11 = 0
Else
    If G11 = "��Ҫ" Then
        H11 = Round(600 * G4 * G7 / 350, 0)
    Else
        H11 = 0
    End If
End If
If I2 = "ˮ��" Then
    J11 = 0
Else
    If I11 = "��Ҫ" Then
        J11 = Round(600 * I4 * I7 / 350, 0)
    Else
        J11 = 0
    End If
End If
If C2 = "ˮ��" Then
    D12 = 0
Else
    If C12 = "��Ҫ" Then
        D12 = Round(300 * C4 * (1 + C7 / 350), 0)
    Else
        D12 = 0
    End If
End If
If E2 = "ˮ��" Then
    F12 = 0
Else
    If E12 = "��Ҫ" Then
        F12 = Round(300 * E4 * (1 + E7 / 350), 0)
    Else
        F12 = 0
    End If
End If
If G2 = "ˮ��" Then
    H12 = 0
Else
    If G12 = "��Ҫ" Then
        H12 = Round(300 * G4 * (1 + G7 / 350), 0)
    Else
        H12 = 0
    End If
End If
If I2 = "ˮ��" Then
    J12 = 0
Else
    If I12 = "��Ҫ" Then
        J12 = Round(300 * I4 * (1 + I7 / 350), 0)
    Else
        J12 = 0
    End If
End If
If C13 = "��Ҫ" Then
    D13 = Round(400 * (1 + C7 / 1050), 0)
Else
    D13 = 0
End If
If E13 = "��Ҫ" Then
    F13 = Round(400 * (1 + E7 / 1050), 0)
Else
    F13 = 0
End If
If G13 = "��Ҫ" Then
    H13 = Round(400 * (1 + G7 / 1050), 0)
Else
    H13 = 0
End If
If I13 = "��Ҫ" Then
    J13 = Round(400 * (1 + I7 / 1050), 0)
Else
    J13 = 0
End If
If C14 = "��Ҫ" Then
    D14 = 600 * C4
Else
    D14 = 0
End If
If E14 = "��Ҫ" Then
    F14 = 600 * E4
Else
    F14 = 0
End If
If G14 = "��Ҫ" Then
    H14 = 600 * G4
Else
    H14 = 0
End If
If I14 = "��Ҫ" Then
    J14 = 600 * I4
Else
    J14 = 0
End If
If C15 = "��Ҫ" Then
    D15 = 600 * C4
Else
    D15 = 0
End If
If E15 = "��Ҫ" Then
    F15 = 600 * E4
Else
    F15 = 0
End If
If G15 = "��Ҫ" Then
    H15 = 600 * G4
Else
    H15 = 0
End If
If I15 = "��Ҫ" Then
    J15 = 600 * I4
Else
    J15 = 0
End If
If C16 = "��Ҫ" Then
    D16 = 400 * C4
Else
    D16 = 0
End If
If E16 = "��Ҫ" Then
    F16 = 400 * E4
Else
    F16 = 0
End If
If G16 = "��Ҫ" Then
    H16 = 400 * G4
Else
    H16 = 0
End If
If I16 = "��Ҫ" Then
    J16 = 400 * I4
Else
    J16 = 0
End If
If C17 = "��Ҫ" Then
    D17 = 400 * C4
Else
    D17 = 0
End If
If E17 = "��Ҫ" Then
    F17 = 400 * E4
Else
    F17 = 0
End If
If G17 = "��Ҫ" Then
    H17 = 400 * G4
Else
    H17 = 0
End If
If I17 = "��Ҫ" Then
    J17 = 400 * I4
Else
    J17 = 0
End If
D18 = Round(1200 * (Val(C4) + Val(E4) + Val(G4) + Val(I4)) * (1 + ((Val(C7) * Val(C4) + Val(E7) * Val(E4) + Val(G7) * Val(G4) + Val(I7) * Val(I4)) / (Val(C4) + Val(E4) + Val(G4) + Val(I4)) - 700) / 3500), 0)
D19 = 350 * (Val(C19) + Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1)
D20 = 600 * (3 + (Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1) * 2)
D21 = (Val(D8) + Val(D9) + Val(D10) + Val(D11) + Val(D12) + Val(D13) + Val(D14) + Val(D15) + Val(D16) + Val(D17)) + _
(Val(F8) + Val(F9) + Val(F10) + Val(F11) + Val(F12) + Val(F13) + Val(F14) + Val(F15) + Val(F16) + Val(F17)) + _
        (Val(H8) + Val(H9) + Val(H10) + Val(H11) + Val(H12) + Val(H13) + Val(H14) + Val(H15) + Val(H16) + Val(H17)) + _
        (Val(J8) + Val(J9) + Val(J10) + Val(J11) + Val(J12) + Val(J13) + Val(J14) + Val(J15) + Val(J16) + Val(J17)) + Val(D18) + Val(D19) + Val(D20)
D22 = Int(D21 * 1.5)
End Sub


Private Sub Form_Load()
Me.Height = 8685
Me.Width = 14340
End Sub

