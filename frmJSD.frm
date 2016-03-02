VERSION 5.00
Begin VB.Form frmWBXT 
   BackColor       =   &H00C0FFC0&
   Caption         =   "速达金额"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14280
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   14280
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   13740
      Picture         =   "frmJSD.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "返回"
      Top             =   7830
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   13260
      Picture         =   "frmJSD.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "保存"
      Top             =   7860
      Width           =   465
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   12750
      Picture         =   "frmJSD.frx":076C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "修改"
      Top             =   7860
      Width           =   465
   End
   Begin VB.Frame Frame1 
      Caption         =   "主机维保"
      Height          =   8475
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   14235
      Begin VB.TextBox txtCJR 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   10740
         TabIndex        =   170
         Top             =   6930
         Width           =   1155
      End
      Begin VB.TextBox txtBz 
         Height          =   7125
         Left            =   12270
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   168
         Text            =   "frmJSD.frx":0A76
         Top             =   570
         Width           =   1815
      End
      Begin VB.ComboBox C8 
         Height          =   300
         ItemData        =   "frmJSD.frx":0A7C
         Left            =   4050
         List            =   "frmJSD.frx":0A86
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   166
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Timer timWait 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   360
         Top             =   570
      End
      Begin VB.Timer timQuit 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   2010
      End
      Begin VB.ComboBox D22 
         ForeColor       =   &H80000001&
         Height          =   300
         ItemData        =   "frmJSD.frx":0A98
         Left            =   4020
         List            =   "frmJSD.frx":0AA2
         TabIndex        =   163
         Text            =   "Combo1"
         Top             =   7170
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   2235
         Left            =   6060
         TabIndex        =   4
         Top             =   6030
         Width           =   4485
         Begin VB.ComboBox Q3 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            ItemData        =   "frmJSD.frx":0AB4
            Left            =   3630
            List            =   "frmJSD.frx":0AC1
            Style           =   2  'Dropdown List
            TabIndex        =   161
            Top             =   630
            Width           =   825
         End
         Begin VB.ComboBox P3 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            ItemData        =   "frmJSD.frx":0AD4
            Left            =   2820
            List            =   "frmJSD.frx":0AE1
            Style           =   2  'Dropdown List
            TabIndex        =   160
            Top             =   630
            Width           =   825
         End
         Begin VB.ComboBox O3 
            BackColor       =   &H00C0FFFF&
            Height          =   300
            ItemData        =   "frmJSD.frx":0AF4
            Left            =   2010
            List            =   "frmJSD.frx":0B01
            Style           =   2  'Dropdown List
            TabIndex        =   159
            Top             =   630
            Width           =   825
         End
         Begin VB.TextBox N2 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1140
            TabIndex        =   26
            Top             =   360
            Width           =   795
         End
         Begin VB.TextBox N3 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   630
            Width           =   795
         End
         Begin VB.TextBox N4 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   900
            Width           =   795
         End
         Begin VB.TextBox N5 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1140
            Width           =   795
         End
         Begin VB.TextBox N6 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   22
            Text            =   "Text5"
            Top             =   1380
            Width           =   795
         End
         Begin VB.TextBox N7 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   21
            Text            =   "Text6"
            Top             =   1650
            Width           =   795
         End
         Begin VB.TextBox O2 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   2010
            TabIndex        =   20
            Text            =   "20-40"
            Top             =   360
            Width           =   795
         End
         Begin VB.TextBox O4 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "Text3"
            Top             =   900
            Width           =   795
         End
         Begin VB.TextBox O5 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   18
            Text            =   "Text4"
            Top             =   1140
            Width           =   795
         End
         Begin VB.TextBox O6 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   17
            Text            =   "Text5"
            Top             =   1380
            Width           =   795
         End
         Begin VB.TextBox O7 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   2010
            Locked          =   -1  'True
            TabIndex        =   16
            Text            =   "Text6"
            Top             =   1650
            Width           =   795
         End
         Begin VB.TextBox P2 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   2820
            TabIndex        =   15
            Text            =   "41-260"
            Top             =   360
            Width           =   795
         End
         Begin VB.TextBox P4 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   2820
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "Text3"
            Top             =   900
            Width           =   795
         End
         Begin VB.TextBox P5 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   2820
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "Text4"
            Top             =   1140
            Width           =   795
         End
         Begin VB.TextBox P6 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   2820
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "Text5"
            Top             =   1380
            Width           =   795
         End
         Begin VB.TextBox P7 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   2820
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "Text6"
            Top             =   1650
            Width           =   795
         End
         Begin VB.TextBox Q2 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   3630
            TabIndex        =   10
            Text            =   "261以上"
            Top             =   360
            Width           =   795
         End
         Begin VB.TextBox Q4 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   3630
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "Text3"
            Top             =   900
            Width           =   795
         End
         Begin VB.TextBox Q5 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   3630
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "Text4"
            Top             =   1140
            Width           =   795
         End
         Begin VB.TextBox Q6 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   3630
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "Text5"
            Top             =   1380
            Width           =   795
         End
         Begin VB.TextBox Q7 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   3630
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "Text6"
            Top             =   1650
            Width           =   795
         End
         Begin VB.TextBox N8 
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "Text25"
            Top             =   1920
            Width           =   3285
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "与内环距离"
            Height          =   255
            Left            =   60
            TabIndex        =   35
            Top             =   390
            Width           =   975
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "单次金额"
            Height          =   255
            Left            =   60
            TabIndex        =   34
            Top             =   690
            Width           =   975
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "年保"
            Height          =   255
            Left            =   60
            TabIndex        =   33
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "巡视"
            Height          =   255
            Left            =   60
            TabIndex        =   32
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "急修"
            Height          =   165
            Left            =   60
            TabIndex        =   31
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "小计"
            Height          =   165
            Left            =   60
            TabIndex        =   30
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "差旅费合计"
            Height          =   255
            Left            =   60
            TabIndex        =   29
            Top             =   1950
            Width           =   975
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "交通费"
            Height          =   195
            Left            =   1170
            TabIndex        =   28
            Top             =   150
            Width           =   945
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "住宿费(按距离公里分段)"
            Height          =   225
            Left            =   2280
            TabIndex        =   27
            Top             =   150
            Width           =   1995
         End
      End
      Begin VB.ComboBox C2 
         Height          =   300
         ItemData        =   "frmJSD.frx":0B14
         Left            =   4050
         List            =   "frmJSD.frx":0B21
         Style           =   2  'Dropdown List
         TabIndex        =   134
         Top             =   240
         Width           =   1785
      End
      Begin VB.ComboBox E2 
         Height          =   300
         ItemData        =   "frmJSD.frx":0B3B
         Left            =   6135
         List            =   "frmJSD.frx":0B48
         Style           =   2  'Dropdown List
         TabIndex        =   133
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox G2 
         Height          =   300
         ItemData        =   "frmJSD.frx":0B62
         Left            =   8250
         List            =   "frmJSD.frx":0B6F
         Style           =   2  'Dropdown List
         TabIndex        =   132
         Top             =   240
         Width           =   1845
      End
      Begin VB.ComboBox I2 
         Height          =   300
         ItemData        =   "frmJSD.frx":0B89
         Left            =   10290
         List            =   "frmJSD.frx":0B96
         Style           =   2  'Dropdown List
         TabIndex        =   131
         Top             =   240
         Width           =   1785
      End
      Begin VB.ComboBox C4 
         Height          =   300
         ItemData        =   "frmJSD.frx":0BB0
         Left            =   4050
         List            =   "frmJSD.frx":0BE7
         Style           =   2  'Dropdown List
         TabIndex        =   130
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox E4 
         Height          =   300
         ItemData        =   "frmJSD.frx":0C25
         Left            =   6135
         List            =   "frmJSD.frx":0C5C
         Style           =   2  'Dropdown List
         TabIndex        =   129
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox G4 
         Height          =   300
         ItemData        =   "frmJSD.frx":0C9A
         Left            =   8235
         List            =   "frmJSD.frx":0CD1
         Style           =   2  'Dropdown List
         TabIndex        =   128
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox I4 
         Height          =   300
         ItemData        =   "frmJSD.frx":0D0F
         Left            =   10290
         List            =   "frmJSD.frx":0D46
         Style           =   2  'Dropdown List
         TabIndex        =   127
         Top             =   960
         Width           =   1005
      End
      Begin VB.ComboBox C5 
         Height          =   300
         ItemData        =   "frmJSD.frx":0D84
         Left            =   4050
         List            =   "frmJSD.frx":0DA3
         Style           =   2  'Dropdown List
         TabIndex        =   126
         Top             =   1320
         Width           =   1005
      End
      Begin VB.ComboBox E5 
         Height          =   300
         ItemData        =   "frmJSD.frx":0DC2
         Left            =   6135
         List            =   "frmJSD.frx":0DE1
         Style           =   2  'Dropdown List
         TabIndex        =   125
         Top             =   1320
         Width           =   1005
      End
      Begin VB.ComboBox G5 
         Height          =   300
         ItemData        =   "frmJSD.frx":0E00
         Left            =   8235
         List            =   "frmJSD.frx":0E1F
         Style           =   2  'Dropdown List
         TabIndex        =   124
         Top             =   1320
         Width           =   1005
      End
      Begin VB.ComboBox I5 
         Height          =   300
         ItemData        =   "frmJSD.frx":0E3E
         Left            =   10290
         List            =   "frmJSD.frx":0E5D
         Style           =   2  'Dropdown List
         TabIndex        =   123
         Top             =   1320
         Width           =   1005
      End
      Begin VB.ComboBox E6 
         Height          =   300
         ItemData        =   "frmJSD.frx":0E7C
         Left            =   6135
         List            =   "frmJSD.frx":0EAD
         TabIndex        =   122
         Text            =   "E6"
         Top             =   1680
         Width           =   1005
      End
      Begin VB.ComboBox G6 
         Height          =   300
         ItemData        =   "frmJSD.frx":0F02
         Left            =   8235
         List            =   "frmJSD.frx":0F33
         TabIndex        =   121
         Text            =   "G6"
         Top             =   1680
         Width           =   1005
      End
      Begin VB.ComboBox I6 
         Height          =   300
         ItemData        =   "frmJSD.frx":0F88
         Left            =   10290
         List            =   "frmJSD.frx":0FB9
         TabIndex        =   120
         Text            =   "I6"
         Top             =   1680
         Width           =   1005
      End
      Begin VB.TextBox C7 
         Height          =   300
         Left            =   4050
         Locked          =   -1  'True
         TabIndex        =   119
         Top             =   2040
         Width           =   1005
      End
      Begin VB.TextBox E7 
         Height          =   300
         Left            =   6135
         Locked          =   -1  'True
         TabIndex        =   118
         Top             =   2040
         Width           =   1005
      End
      Begin VB.TextBox G7 
         Height          =   300
         Left            =   8235
         Locked          =   -1  'True
         TabIndex        =   117
         Top             =   2040
         Width           =   1005
      End
      Begin VB.TextBox I7 
         Height          =   300
         Left            =   10290
         Locked          =   -1  'True
         TabIndex        =   116
         Top             =   2040
         Width           =   1005
      End
      Begin VB.ComboBox C6 
         Height          =   300
         ItemData        =   "frmJSD.frx":100E
         Left            =   4050
         List            =   "frmJSD.frx":103F
         TabIndex        =   115
         Text            =   "C6"
         Top             =   1680
         Width           =   1005
      End
      Begin VB.ComboBox C10 
         Height          =   300
         ItemData        =   "frmJSD.frx":1094
         Left            =   4050
         List            =   "frmJSD.frx":10A1
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   3150
         Width           =   1005
      End
      Begin VB.ComboBox C11 
         Height          =   300
         ItemData        =   "frmJSD.frx":10BD
         Left            =   4050
         List            =   "frmJSD.frx":10C7
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   113
         Top             =   3510
         Width           =   1005
      End
      Begin VB.ComboBox C12 
         Height          =   300
         ItemData        =   "frmJSD.frx":10D9
         Left            =   4050
         List            =   "frmJSD.frx":10E3
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   3885
         Width           =   1005
      End
      Begin VB.ComboBox C13 
         Height          =   300
         ItemData        =   "frmJSD.frx":10F5
         Left            =   4050
         List            =   "frmJSD.frx":10FF
         Style           =   2  'Dropdown List
         TabIndex        =   111
         Top             =   4245
         Width           =   1005
      End
      Begin VB.ComboBox C14 
         Height          =   300
         ItemData        =   "frmJSD.frx":1111
         Left            =   4050
         List            =   "frmJSD.frx":111B
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   4620
         Width           =   1005
      End
      Begin VB.ComboBox C15 
         Height          =   300
         ItemData        =   "frmJSD.frx":112D
         Left            =   4050
         List            =   "frmJSD.frx":1137
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   4995
         Width           =   1005
      End
      Begin VB.ComboBox C16 
         Height          =   300
         ItemData        =   "frmJSD.frx":1149
         Left            =   4050
         List            =   "frmJSD.frx":1153
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   5355
         Width           =   1005
      End
      Begin VB.ComboBox C17 
         Height          =   300
         ItemData        =   "frmJSD.frx":1165
         Left            =   4050
         List            =   "frmJSD.frx":116F
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   5730
         Width           =   1005
      End
      Begin VB.ComboBox E10 
         Height          =   300
         ItemData        =   "frmJSD.frx":1181
         Left            =   6135
         List            =   "frmJSD.frx":118E
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   3150
         Width           =   1005
      End
      Begin VB.ComboBox E11 
         Height          =   300
         ItemData        =   "frmJSD.frx":11AA
         Left            =   6135
         List            =   "frmJSD.frx":11B4
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   3510
         Width           =   1005
      End
      Begin VB.ComboBox E12 
         Height          =   300
         ItemData        =   "frmJSD.frx":11C6
         Left            =   6135
         List            =   "frmJSD.frx":11D0
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   3885
         Width           =   1005
      End
      Begin VB.ComboBox E13 
         Height          =   300
         ItemData        =   "frmJSD.frx":11E2
         Left            =   6135
         List            =   "frmJSD.frx":11EC
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   4245
         Width           =   1005
      End
      Begin VB.ComboBox E14 
         Height          =   300
         ItemData        =   "frmJSD.frx":11FE
         Left            =   6135
         List            =   "frmJSD.frx":1208
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   4620
         Width           =   1005
      End
      Begin VB.ComboBox E15 
         Height          =   300
         ItemData        =   "frmJSD.frx":121A
         Left            =   6135
         List            =   "frmJSD.frx":1224
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   4995
         Width           =   1005
      End
      Begin VB.ComboBox E16 
         Height          =   300
         ItemData        =   "frmJSD.frx":1236
         Left            =   6135
         List            =   "frmJSD.frx":1240
         Style           =   2  'Dropdown List
         TabIndex        =   100
         Top             =   5355
         Width           =   1005
      End
      Begin VB.ComboBox E17 
         Height          =   300
         ItemData        =   "frmJSD.frx":1252
         Left            =   6135
         List            =   "frmJSD.frx":125C
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   5730
         Width           =   1005
      End
      Begin VB.ComboBox G10 
         Height          =   300
         ItemData        =   "frmJSD.frx":126E
         Left            =   8235
         List            =   "frmJSD.frx":127B
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   3150
         Width           =   1005
      End
      Begin VB.ComboBox G11 
         Height          =   300
         ItemData        =   "frmJSD.frx":1297
         Left            =   8235
         List            =   "frmJSD.frx":12A1
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   3510
         Width           =   1005
      End
      Begin VB.ComboBox G12 
         Height          =   300
         ItemData        =   "frmJSD.frx":12B3
         Left            =   8235
         List            =   "frmJSD.frx":12BD
         Style           =   2  'Dropdown List
         TabIndex        =   96
         Top             =   3885
         Width           =   1005
      End
      Begin VB.ComboBox G13 
         Height          =   300
         ItemData        =   "frmJSD.frx":12CF
         Left            =   8235
         List            =   "frmJSD.frx":12D9
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   4245
         Width           =   1005
      End
      Begin VB.ComboBox G14 
         Height          =   300
         ItemData        =   "frmJSD.frx":12EB
         Left            =   8235
         List            =   "frmJSD.frx":12F5
         Style           =   2  'Dropdown List
         TabIndex        =   94
         Top             =   4620
         Width           =   1005
      End
      Begin VB.ComboBox G15 
         Height          =   300
         ItemData        =   "frmJSD.frx":1307
         Left            =   8235
         List            =   "frmJSD.frx":1311
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   4995
         Width           =   1005
      End
      Begin VB.ComboBox G16 
         Height          =   300
         ItemData        =   "frmJSD.frx":1323
         Left            =   8235
         List            =   "frmJSD.frx":132D
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   5355
         Width           =   1005
      End
      Begin VB.ComboBox G17 
         Height          =   300
         ItemData        =   "frmJSD.frx":133F
         Left            =   8235
         List            =   "frmJSD.frx":1349
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   5730
         Width           =   1005
      End
      Begin VB.ComboBox I9 
         Height          =   300
         ItemData        =   "frmJSD.frx":135B
         Left            =   10290
         List            =   "frmJSD.frx":1368
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   2775
         Width           =   1005
      End
      Begin VB.ComboBox I11 
         Height          =   300
         ItemData        =   "frmJSD.frx":1384
         Left            =   10290
         List            =   "frmJSD.frx":138E
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   3513
         Width           =   1005
      End
      Begin VB.ComboBox I12 
         Height          =   300
         ItemData        =   "frmJSD.frx":13A0
         Left            =   10290
         List            =   "frmJSD.frx":13AA
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   3882
         Width           =   1005
      End
      Begin VB.ComboBox I13 
         Height          =   300
         ItemData        =   "frmJSD.frx":13BC
         Left            =   10290
         List            =   "frmJSD.frx":13C6
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   4251
         Width           =   1005
      End
      Begin VB.ComboBox I14 
         Height          =   300
         ItemData        =   "frmJSD.frx":13D8
         Left            =   10290
         List            =   "frmJSD.frx":13E2
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   4620
         Width           =   1005
      End
      Begin VB.ComboBox I15 
         Height          =   300
         ItemData        =   "frmJSD.frx":13F4
         Left            =   10290
         List            =   "frmJSD.frx":13FE
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   4989
         Width           =   1005
      End
      Begin VB.ComboBox I16 
         Height          =   300
         ItemData        =   "frmJSD.frx":1410
         Left            =   10290
         List            =   "frmJSD.frx":141A
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   5358
         Width           =   1005
      End
      Begin VB.ComboBox I17 
         Height          =   300
         ItemData        =   "frmJSD.frx":142C
         Left            =   10290
         List            =   "frmJSD.frx":1436
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   5730
         Width           =   1005
      End
      Begin VB.ComboBox C9 
         Height          =   300
         ItemData        =   "frmJSD.frx":1448
         Left            =   4050
         List            =   "frmJSD.frx":1455
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   2775
         Width           =   1005
      End
      Begin VB.ComboBox E9 
         Height          =   300
         ItemData        =   "frmJSD.frx":1471
         Left            =   6135
         List            =   "frmJSD.frx":147E
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   2775
         Width           =   1005
      End
      Begin VB.ComboBox G9 
         Height          =   300
         ItemData        =   "frmJSD.frx":149A
         Left            =   8235
         List            =   "frmJSD.frx":14A7
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   2775
         Width           =   1005
      End
      Begin VB.TextBox D8 
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   79
         Text            =   "Text5"
         Top             =   2400
         Width           =   885
      End
      Begin VB.TextBox D9 
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "Text6"
         Top             =   2775
         Width           =   885
      End
      Begin VB.TextBox D10 
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   77
         Text            =   "Text5"
         Top             =   3157
         Width           =   885
      End
      Begin VB.TextBox D11 
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   76
         Text            =   "Text5"
         Top             =   3524
         Width           =   885
      End
      Begin VB.TextBox D12 
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   75
         Text            =   "Text5"
         Top             =   3891
         Width           =   885
      End
      Begin VB.TextBox D13 
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   74
         Text            =   "Text5"
         Top             =   4258
         Width           =   885
      End
      Begin VB.TextBox D14 
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   73
         Text            =   "Text5"
         Top             =   4625
         Width           =   885
      End
      Begin VB.TextBox D15 
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   72
         Text            =   "Text5"
         Top             =   4992
         Width           =   885
      End
      Begin VB.TextBox D16 
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "Text5"
         Top             =   5359
         Width           =   885
      End
      Begin VB.TextBox D17 
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "Text5"
         Top             =   5730
         Width           =   885
      End
      Begin VB.TextBox F9 
         Height          =   270
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "Text6"
         Top             =   2775
         Width           =   855
      End
      Begin VB.TextBox F10 
         Height          =   270
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "Text5"
         Top             =   3150
         Width           =   855
      End
      Begin VB.TextBox F11 
         Height          =   270
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "Text5"
         Top             =   3525
         Width           =   855
      End
      Begin VB.TextBox F12 
         Height          =   270
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   66
         Text            =   "Text5"
         Top             =   3885
         Width           =   855
      End
      Begin VB.TextBox F13 
         Height          =   270
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "Text5"
         Top             =   4260
         Width           =   855
      End
      Begin VB.TextBox F14 
         Height          =   270
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Text5"
         Top             =   4620
         Width           =   855
      End
      Begin VB.TextBox F15 
         Height          =   270
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "Text5"
         Top             =   4995
         Width           =   855
      End
      Begin VB.TextBox F16 
         Height          =   270
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "Text5"
         Top             =   5355
         Width           =   855
      End
      Begin VB.TextBox F17 
         Height          =   270
         Left            =   7230
         Locked          =   -1  'True
         TabIndex        =   61
         Text            =   "Text5"
         Top             =   5730
         Width           =   855
      End
      Begin VB.TextBox H9 
         Height          =   270
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   "Text6"
         Top             =   2775
         Width           =   825
      End
      Begin VB.TextBox H11 
         Height          =   270
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   59
         Text            =   "Text5"
         Top             =   3525
         Width           =   825
      End
      Begin VB.TextBox H12 
         Height          =   270
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Text5"
         Top             =   3885
         Width           =   825
      End
      Begin VB.TextBox H13 
         Height          =   270
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "Text5"
         Top             =   4260
         Width           =   825
      End
      Begin VB.TextBox H14 
         Height          =   270
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "Text5"
         Top             =   4620
         Width           =   825
      End
      Begin VB.TextBox H15 
         Height          =   270
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "Text5"
         Top             =   4995
         Width           =   825
      End
      Begin VB.TextBox H16 
         Height          =   270
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   54
         Text            =   "Text5"
         Top             =   5355
         Width           =   825
      End
      Begin VB.TextBox H17 
         Height          =   270
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "Text5"
         Top             =   5730
         Width           =   825
      End
      Begin VB.TextBox J9 
         Height          =   270
         Left            =   11370
         Locked          =   -1  'True
         TabIndex        =   52
         Text            =   "Text6"
         Top             =   2775
         Width           =   825
      End
      Begin VB.TextBox J11 
         Height          =   270
         Left            =   11370
         Locked          =   -1  'True
         TabIndex        =   51
         Text            =   "Text5"
         Top             =   3513
         Width           =   825
      End
      Begin VB.TextBox J12 
         Height          =   270
         Left            =   11370
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "Text5"
         Top             =   3882
         Width           =   825
      End
      Begin VB.TextBox J13 
         Height          =   270
         Left            =   11370
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "Text5"
         Top             =   4251
         Width           =   825
      End
      Begin VB.TextBox J14 
         Height          =   270
         Left            =   11370
         Locked          =   -1  'True
         TabIndex        =   48
         Text            =   "Text5"
         Top             =   4620
         Width           =   825
      End
      Begin VB.TextBox J15 
         Height          =   270
         Left            =   11370
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "Text5"
         Top             =   4989
         Width           =   825
      End
      Begin VB.TextBox J16 
         Height          =   270
         Left            =   11370
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "Text5"
         Top             =   5358
         Width           =   825
      End
      Begin VB.TextBox J17 
         Height          =   270
         Left            =   11370
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "Text5"
         Top             =   5730
         Width           =   825
      End
      Begin VB.TextBox D18 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "Text5"
         Top             =   6105
         Width           =   885
      End
      Begin VB.TextBox D19 
         ForeColor       =   &H00FF00FF&
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "Text5"
         Top             =   6450
         Width           =   885
      End
      Begin VB.TextBox D20 
         ForeColor       =   &H000080FF&
         Height          =   270
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "Text5"
         Top             =   6795
         Width           =   885
      End
      Begin VB.TextBox D21 
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Text5"
         Top             =   7485
         Width           =   2025
      End
      Begin VB.ComboBox I10 
         Height          =   300
         ItemData        =   "frmJSD.frx":14C3
         Left            =   10290
         List            =   "frmJSD.frx":14D0
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   3144
         Width           =   1005
      End
      Begin VB.TextBox H10 
         Height          =   270
         Left            =   9300
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Text5"
         Top             =   3150
         Width           =   825
      End
      Begin VB.TextBox J10 
         Height          =   270
         Left            =   11370
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Text5"
         Top             =   3144
         Width           =   825
      End
      Begin VB.CommandButton cmdJi 
         BackColor       =   &H00FF8080&
         Caption         =   "计算"
         Height          =   375
         Left            =   10590
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   7860
         Width           =   1695
      End
      Begin VB.ComboBox C19 
         Height          =   300
         ItemData        =   "frmJSD.frx":14EC
         Left            =   4050
         List            =   "frmJSD.frx":1508
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   6450
         Width           =   1005
      End
      Begin VB.Label Label36 
         Caption         =   "承接人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   10770
         TabIndex        =   169
         Top             =   6540
         Width           =   795
      End
      Begin VB.Label Label25 
         Caption         =   "备注"
         Height          =   225
         Left            =   12300
         TabIndex        =   167
         Top             =   300
         Width           =   525
      End
      Begin VB.Label lblBid 
         Caption         =   "lblBid"
         Height          =   195
         Left            =   450
         TabIndex        =   165
         Top             =   4800
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblMid 
         Caption         =   "lblMid"
         Height          =   315
         Left            =   540
         TabIndex        =   164
         Top             =   5130
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label35 
         Caption         =   "全包"
         ForeColor       =   &H80000001&
         Height          =   255
         Left            =   2400
         TabIndex        =   162
         Top             =   7200
         Width           =   1005
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   6060
         X2              =   6060
         Y1              =   90
         Y2              =   8310
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   8160
         X2              =   8160
         Y1              =   90
         Y2              =   6120
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   10200
         X2              =   10200
         Y1              =   90
         Y2              =   7770
      End
      Begin VB.Line Line6 
         X1              =   10530
         X2              =   -960
         Y1              =   7770
         Y2              =   7770
      End
      Begin VB.Label Label1 
         Caption         =   "类型"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   158
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "内容"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2070
         TabIndex        =   157
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "询价参数"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   300
         TabIndex        =   156
         Top             =   1050
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "年度保养"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   210
         TabIndex        =   155
         Top             =   3420
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmJSD.frx":152A
         Height          =   225
         Left            =   1350
         TabIndex        =   154
         Top             =   6150
         Width           =   2085
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "机组数量"
         Height          =   285
         Left            =   1170
         TabIndex        =   153
         Top             =   990
         Width           =   2265
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "每台机组压缩机数量"
         Height          =   285
         Left            =   1170
         TabIndex        =   152
         Top             =   1365
         Width           =   2265
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmJSD.frx":1540
         Height          =   285
         Left            =   1170
         TabIndex        =   151
         Top             =   1725
         Width           =   2265
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmJSD.frx":155A
         Height          =   285
         Left            =   1170
         TabIndex        =   150
         Top             =   2100
         Width           =   2265
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmJSD.frx":156A
         Height          =   285
         Left            =   1170
         TabIndex        =   149
         Top             =   2460
         Width           =   2265
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "冷凝器机械清洗"
         Height          =   285
         Left            =   1170
         TabIndex        =   148
         Top             =   2835
         Width           =   2265
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmJSD.frx":157C
         Height          =   285
         Left            =   1170
         TabIndex        =   147
         Top             =   3195
         Width           =   2265
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "风扇电机保养"
         Height          =   285
         Left            =   1170
         TabIndex        =   146
         Top             =   3570
         Width           =   2265
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmJSD.frx":1590
         Height          =   285
         Left            =   1170
         TabIndex        =   145
         Top             =   3930
         Width           =   2265
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmJSD.frx":15A4
         Height          =   285
         Left            =   1170
         TabIndex        =   144
         Top             =   4305
         Width           =   2265
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "换冷冻油及过滤器"
         Height          =   285
         Left            =   1170
         TabIndex        =   143
         Top             =   4665
         Width           =   2265
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmJSD.frx":15B2
         Height          =   285
         Left            =   1170
         TabIndex        =   142
         Top             =   5040
         Width           =   2265
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmJSD.frx":15C8
         Height          =   285
         Left            =   1170
         TabIndex        =   141
         Top             =   5400
         Width           =   2265
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmJSD.frx":15D6
         Height          =   285
         Left            =   1170
         TabIndex        =   140
         Top             =   5775
         Width           =   2265
      End
      Begin VB.Label Label20 
         Caption         =   "选项         金额       选项         金额     选项          金额     选项         金额 "
         Height          =   255
         Left            =   4110
         TabIndex        =   139
         Top             =   600
         Width           =   8055
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   3720
         X2              =   3720
         Y1              =   90
         Y2              =   8310
      End
      Begin VB.Label Label21 
         Caption         =   "调试"
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   870
         TabIndex        =   138
         Top             =   6180
         Width           =   585
      End
      Begin VB.Label Label22 
         Caption         =   "巡视"
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   870
         TabIndex        =   137
         Top             =   6465
         Width           =   405
      End
      Begin VB.Label Label23 
         Caption         =   "急修"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   870
         TabIndex        =   136
         Top             =   6795
         Width           =   465
      End
      Begin VB.Line Line2 
         X1              =   6270
         X2              =   -960
         Y1              =   6390
         Y2              =   6390
      End
      Begin VB.Line Line3 
         X1              =   6270
         X2              =   -960
         Y1              =   6750
         Y2              =   6750
      End
      Begin VB.Line Line4 
         X1              =   10470
         X2              =   -1170
         Y1              =   7140
         Y2              =   7140
      End
      Begin VB.Line Line5 
         X1              =   10500
         X2              =   -960
         Y1              =   7470
         Y2              =   7470
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   $"frmJSD.frx":15E0
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1350
         TabIndex        =   135
         Top             =   7560
         Width           =   2115
      End
   End
End
Attribute VB_Name = "frmWBXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer '(1保存 )
Public Sub Qing()
C2.Text = "水冷": E2.Text = "水冷": G2.Text = "水冷": I2.Text = "水冷"
C4.Text = 0: E4.Text = 0: G4.Text = 0: I4.Text = 0
C5.Text = 0: E5.Text = 0: G5.Text = 0: I5.Text = 0
C6.Text = 0: E6.Text = 0: G6.Text = 0: I6.Text = 0
C7.Text = "": E7.Text = "": G7.Text = "": I7.Text = ""
C8.Text = "需要"
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
D18.Text = "": D19.Text = "": D20.Text = "": D21.Text = "": D22.Text = "非全包"
C9.Text = "不需要": E9.Text = "不需要": G9.Text = "不需要": I9.Text = "不需要"
C10.Text = "不需要": E10.Text = "不需要": G10.Text = "不需要": I10.Text = "不需要"
C11.Text = "需要": E11.Text = "需要": G11.Text = "需要": I11.Text = "需要"
C12.Text = "不需要": E12.Text = "不需要": G12.Text = "不需要": I12.Text = "不需要"
C13.Text = "不需要": E13.Text = "不需要": G13.Text = "不需要": I13.Text = "不需要"
C14.Text = "不需要": E14.Text = "不需要": G14.Text = "不需要": I14.Text = "不需要"
C15.Text = "需要": E15.Text = "需要": G15.Text = "需要": I15.Text = "需要"
C16.Text = "不需要": E16.Text = "不需要": G16.Text = "不需要": I16.Text = "不需要"
C17.Text = "需要": E17.Text = "需要": G17.Text = "需要": I17.Text = "需要"
C19.Text = 5
N2.Text = "": N3.Text = "": N4.Text = "": N5.Text = "": N6.Text = "": N7.Text = "": N8.Text = ""
 O4.Text = "": O5.Text = "": O6.Text = "": O7.Text = ""
 P4.Text = "": P5.Text = "": P6.Text = "": P7.Text = ""
 Q4.Text = "": Q5.Text = "": Q6.Text = "": Q7.Text = ""
 lblMid.Caption = ""
 lblBid.Caption = ""
 txtBz.Text = ""
 txtCJR.Text = "" '承接人
 cmdSave.Enabled = False

End Sub

Private Sub C19_LostFocus()
If Me.Visible = False Then Exit Sub
If Val(C19.Text) > 0 Then
    MsgBox "请在备注中写明巡视的大致日期信息!"
End If


End Sub


Private Sub cmdBack_Click()
Me.Visible = False
End Sub



Private Sub cmdJi_Click()
If Me.Visible = False Then Exit Sub
If Val(C4.Text) = 0 And Val(E4.Text) = 0 And Val(G4.Text) = 0 And Val(I4.Text) = 0 Then
    MsgBox "请确定机组数量!"
    Exit Sub
End If
Call J1
End Sub


Private Sub cmdMod_Click()
If Val(frmWBXX.lblLc.Caption) > 1 And mod1.DName <> "" Then
    Exit Sub
End If
cmdSave.Enabled = True
If mod1.DName = "" Then
    C8.Locked = False
    C11.Locked = False
    C15.Locked = False
    C17.Locked = False
End If
End Sub

Private Sub cmdSave_Click()
On Error Resume Next

Call cmdJi_Click

 '保存
    timZm = 1
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "人工业务"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = lblMid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = txtCJR.Text '承接人
    mod1.cmd.Parameters("@mt6") = C2.Text
    mod1.cmd.Parameters("@mt7") = E2.Text
    mod1.cmd.Parameters("@mt8") = G2.Text
    mod1.cmd.Parameters("@mt9") = I2.Text
    mod1.cmd.Parameters("@mt10") = C8.Text
    mod1.cmd.Parameters("@mt11") = C9.Text
    mod1.cmd.Parameters("@mt12") = C10.Text
    mod1.cmd.Parameters("@mt13") = C11.Text
    mod1.cmd.Parameters("@mt14") = C12.Text
    mod1.cmd.Parameters("@mt15") = C13.Text
    mod1.cmd.Parameters("@mt16") = C14.Text
    mod1.cmd.Parameters("@mt17") = C15.Text
    mod1.cmd.Parameters("@mt18") = C16.Text
    mod1.cmd.Parameters("@mt19") = C17.Text
    mod1.cmd.Parameters("@mt20") = E9.Text
    mod1.cmd.Parameters("@mt21") = E10.Text
    mod1.cmd.Parameters("@mt22") = E11.Text
    mod1.cmd.Parameters("@mt23") = E12.Text
    mod1.cmd.Parameters("@mt24") = E13.Text
    mod1.cmd.Parameters("@mt25") = E14.Text
    mod1.cmd.Parameters("@mt26") = E15.Text
    mod1.cmd.Parameters("@mt27") = E16.Text
    mod1.cmd.Parameters("@mt28") = E17.Text
    mod1.cmd.Parameters("@mt29") = G9.Text
    mod1.cmd.Parameters("@mt30") = G10.Text
    mod1.cmd.Parameters("@mt31") = G11.Text
    mod1.cmd.Parameters("@mt32") = G12.Text
    mod1.cmd.Parameters("@mt33") = G13.Text
    mod1.cmd.Parameters("@mt34") = G14.Text
    mod1.cmd.Parameters("@mt35") = G15.Text
    mod1.cmd.Parameters("@mt36") = G16.Text
    mod1.cmd.Parameters("@mt37") = G17.Text
    mod1.cmd.Parameters("@mt38") = I9.Text
    mod1.cmd.Parameters("@mt39") = I10.Text
    mod1.cmd.Parameters("@mt40") = I11.Text
    mod1.cmd.Parameters("@mt41") = I12.Text
    mod1.cmd.Parameters("@mt42") = I13.Text
    mod1.cmd.Parameters("@mlt1") = txtBz.Text '备注
    mod1.cmd.Parameters("@mlt2") = I14.Text
    mod1.cmd.Parameters("@mlt3") = I15.Text
    mod1.cmd.Parameters("@mlt4") = I16.Text
    mod1.cmd.Parameters("@mlt5") = I17.Text
    mod1.cmd.Parameters("@mm1") = Val(D21.Text) '基准价
    mod1.cmd.Parameters("@mm2") = Val(N8.Text)
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = Val(lblBid.Caption) '询价单号
    mod1.cmd.Parameters("@mm6") = Val(C4.Text)
    mod1.cmd.Parameters("@mm7") = Val(E4.Text)
    mod1.cmd.Parameters("@mm8") = Val(G4.Text)
    mod1.cmd.Parameters("@mm9") = Val(I4.Text)
    mod1.cmd.Parameters("@mm10") = Val(C5.Text)
    mod1.cmd.Parameters("@mm11") = Val(E5.Text)
    mod1.cmd.Parameters("@mm12") = Val(G5.Text)
    mod1.cmd.Parameters("@mm13") = Val(I5.Text)
    mod1.cmd.Parameters("@mm14") = Val(C6.Text)
    mod1.cmd.Parameters("@mm15") = Val(E6.Text)
    mod1.cmd.Parameters("@mm16") = Val(G6.Text)
    mod1.cmd.Parameters("@mm17") = Val(I6.Text)
    mod1.cmd.Parameters("@mm18") = Val(C19.Text) '巡视
    mod1.cmd.Parameters("@mm19") = Val(N2.Text)
    mod1.cmd.Parameters("@mm20") = Val(O3.Text)
    mod1.cmd.Parameters("@mm21") = Val(P3.Text)
    mod1.cmd.Parameters("@mm22") = Val(Q3.Text)
    If D22.Text = "非全包" Then
        mod1.cmd.Parameters("@mb1") = 0
    Else
        mod1.cmd.Parameters("@mb1") = 1
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


    Set mod1.cmd = Nothing

End Sub

Private Sub Combo1_Change()

End Sub

Private Sub E4_Click()
If Me.Visible = False Then
Exit Sub
End If
If E4 = 0 Then
    MsgBox "请从靠前列操作!"
    Exit Sub
End If
If Val(C4.Text) = 0 Then
    E4 = "0"
End If
End Sub

Private Sub Form_Load()
Me.Height = 8685
Me.Width = 14340
Me.Left = 0
Me.Top = 0
End Sub

Private Sub Text19_Change()

End Sub

Public Sub Bound(Mid As Long)
Dim JZ As Single
Dim JT As Single
Dim tt As String
Dim Ra
On Error GoTo frmWBXT5
tt = "select mt1,mt2,mt3,mt4,mt5,mt6,mt7,mt8,mt9,mt10,mt11,mt12,mt13,mt14,mt15,mt16,mt17,mt18,mt19,mt20," & _
    "mt21,mt22,mt23,mt24,mt25,mt26,mt27,mt28,mt29,mt30,mt31,mt32,mt33,mt34,mt35,mt36,mt37,mt38,mt39,mt40,mt41,mt42,mlt1,mlt2,mlt3,mlt4,mlt5," & _
    "mm1,mm2,mm3,mm4,mm5,mm6,mm7,mm8,mm9,mm10,mm11,mm12,mm13,mm14,mm15,mm16,mm17,mm18,mm19,mm20," & _
    "mm21,mm22,mm23,mm24,mm25,mm26,mm27,mm28,mm29,mm30,mb1,mb2,mb3,mb4,mb5,bid,mid" & _
    " from MlMX where mid=" & Mid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next
txtCJR.Text = Ra(4, 0) '承接人
C2.Text = Ra(5, 0)
E2.Text = Ra(6, 0)
G2.Text = Ra(7, 0)
I2.Text = Ra(8, 0)
C8.Text = Ra(9, 0)
C9.Text = Ra(10, 0)
C10.Text = Ra(11, 0)
C11.Text = Ra(12, 0)
C12.Text = Ra(13, 0)
C13.Text = Ra(14, 0)
C14.Text = Ra(15, 0)
C15.Text = Ra(16, 0)
C16.Text = Ra(17, 0)
C17.Text = Ra(18, 0)
C19.Text = Ra(64, 0) '巡视
E9.Text = Ra(19, 0)
E10.Text = Ra(20, 0)
E11.Text = Ra(21, 0)
E12.Text = Ra(22, 0)
E13.Text = Ra(23, 0)
E14.Text = Ra(24, 0)
E15.Text = Ra(25, 0)
E16.Text = Ra(26, 0)
E17.Text = Ra(27, 0)
G9.Text = Ra(28, 0)
G10.Text = Ra(29, 0)
G11.Text = Ra(30, 0)
G12.Text = Ra(31, 0)
G13.Text = Ra(32, 0)
G14.Text = Ra(33, 0)
G15.Text = Ra(34, 0)
G16.Text = Ra(35, 0)
G17.Text = Ra(36, 0)
I9.Text = Ra(37, 0)
I10.Text = Ra(38, 0)
I11.Text = Ra(39, 0)
I12.Text = Ra(40, 0)
I13.Text = Ra(41, 0)
I14.Text = Ra(43, 0)
I15.Text = Ra(44, 0)
I16.Text = Ra(45, 0)
I17.Text = Ra(46, 0)
txtBz.Text = Ra(42, 0) '备注
D22.Text = Ra(47, 0): JZ = Val(D22.Text) '基准价
N8.Text = Ra(48, 0): JT = Val(N8.Text) '差旅
C4.Text = Ra(52, 0)
E4.Text = Ra(53, 0)
G4.Text = Ra(54, 0)
I4.Text = Ra(55, 0)
C5.Text = Ra(56, 0)
E5.Text = Ra(57, 0)
G5.Text = Ra(58, 0)
I5.Text = Ra(59, 0)
C6.Text = Ra(60, 0)
E6.Text = Ra(61, 0)
G6.Text = Ra(62, 0)
I6.Text = Ra(63, 0)
N2.Text = Ra(65, 0)
O3.Text = Ra(66, 0)
P3.Text = Ra(67, 0)
Q3.Text = Ra(68, 0)
If Ra(77, 0) = 0 Then
    D22 = "非全包"
Else
    D22 = "全包"
End If
    

 lblMid.Caption = Ra(83, 0)
 lblBid.Caption = Ra(82, 0)
Call J1
Exit Sub
frmWBXT5:
MsgBox "出错!"
End
End Sub

Private Sub G4_Click()
If Me.Visible = False Then
Exit Sub
End If
If G4 = 0 Then
    MsgBox "请从靠前列操作!"
    Exit Sub
End If
If Val(E4.Text) = 0 Then
    G4 = "0"
End If
End Sub

Private Sub I4_Click()
If Me.Visible = False Then
Exit Sub
End If
If I4 = 0 Then
    MsgBox "请从靠前列操作!"
    Exit Sub
End If
If Val(G4.Text) = 0 Then
    I4 = "0"
End If
End Sub

Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0


If timZm = 1 Then    '保存
    cmdSave.Enabled = False
End If
timQuit.Enabled = False
Me.Enabled = True
End Sub


Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
On Error Resume Next
timWait.Enabled = False
Me.Enabled = False
tt = "select cf,bz,bh,mm1,mm2,mt1,mt2 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then '保存
        Call frmWBXX.MXBound(Val(frmWBXX.lblBid.Caption))
        frmWBXX.txt2.Text = mod1.WP.Fields("mm2").Value
        'Call frmWBXX.ji
    End If

    timWait.Enabled = False

    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdJG.Enabled = False
    End If
    timWait.Enabled = False
    Me.Enabled = True
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdJG.Enabled = False
    End If
    Me.Enabled = True
    Exit Sub

End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub



Public Sub J1()
Dim B2, B3, B4, B5, B6, B7, B9, B10, B11, B12, B13, B14, B15, B16, B17
'先调整,热磊优先
If I2 = "热泵" And G2 <> "热泵" Then
    B2 = G2: B3 = G3: B4 = G4: B5 = G5: B6 = G6: B7 = G7:  B9 = G9: B10 = G10: B11 = G11: B12 = G12: B13 = G13: B14 = G14: B15 = G15: B16 = G16: B17 = G17
    G2 = I2: G3 = I3: G4 = I4: G5 = I5: G6 = I6: G7 = I7:  G9 = I9: G10 = I10: G11 = I11: G12 = I12: G13 = I13: G14 = I14: G15 = I15: G16 = I16: G17 = I17
    I2 = B2: I3 = B3: I4 = B4: I5 = B5: I6 = B6: I7 = B7:  I9 = B9: I10 = B10: I11 = B11: I12 = B12: I13 = B13: I14 = B14: I15 = B15: I16 = B16: I17 = B17
End If
If G2 = "热泵" And E2 <> "热泵" Then
    B2 = E2: B3 = E3: B4 = E4: B5 = E5: B6 = E6: B7 = E7:  B9 = E9: B10 = E10: B11 = E11: B12 = E12: B13 = E13: B14 = E14: B15 = E15: B16 = E16: B17 = E17
    E2 = G2: E3 = G3: E4 = G4: E5 = G5: E6 = G6: E7 = G7:  E9 = G9: E10 = G10: E11 = G11: E12 = G12: E13 = G13: E14 = G14: E15 = G15: E16 = G16: E17 = G17
    G2 = B2: G3 = B3: G4 = B4: G5 = B5: G6 = B6: G7 = B7:  G9 = B9: G10 = B10: G11 = B11: G12 = B12: G13 = B13: G14 = B14: G15 = B15: G16 = B16: G17 = B17
End If
If E2 = "热泵" And C2 <> "热泵" Then
    B2 = C2: B3 = C3: B4 = C4: B5 = C5: B6 = C6: B7 = C7:  B9 = C9: B10 = C10: B11 = C11: B12 = C12: B13 = C13: B14 = C14: B15 = C15: B16 = C16: B17 = C17
    C2 = E2: C3 = E3: C4 = E4: C5 = E5: C6 = E6: C7 = E7:  C9 = E9: C10 = E10: C11 = E11: C12 = E12: C13 = E13: C14 = E14: C15 = E15: C16 = E16: C17 = E17
    E2 = B2: E3 = B3: E4 = B4: E5 = B5: E6 = B6: E7 = B7:  E9 = B9: E10 = B10: E11 = B11: E12 = B12: E13 = B13: E14 = B14: E15 = B15: E16 = B16: E17 = B17
End If


If Val(E4.Text) = 0 Then
    E5.Text = 0
    E6.Text = 0
End If
If Val(G4.Text) = 0 Then
    G5.Text = 0
    G6.Text = 0
End If
If Val(I4.Text) = 0 Then
    I5.Text = 0
    I6.Text = 0
End If
C7 = C6 * C5
E7 = E6 * E5
G7 = G6 * G5
I7 = I6 * I5
If C8.Text = "需要" Then
    D8.Text = 500
Else
    D8.Text = 0
End If
If C9 = "拆一端" Then
    If C2 = "水冷" Then
        D9 = Round(300 * C4 * (1 + C7 / 1050), 0)
    Else
        D9 = 0
    End If
ElseIf C9 = "拆两端" Then
    If C2 = "水冷" Then
        D9 = Round(300 * C4 * (1 + C7 / 1050) * 1.1, 0)
    Else
        D9 = 0
    End If
Else
    D9 = 0
End If

If E9 = "拆一端" Then
    If E2 = "水冷" Then
        F9 = Round(300 * E4 * (1 + E7 / 1050), 0)
    Else
        F9 = 0
    End If
ElseIf E9 = "拆两端" Then
    If E2 = "水冷" Then
        F9 = Round(300 * E4 * (1 + E7 / 1050) * 1.1, 0)
    Else
        F9 = 0
    End If
Else
    F9 = 0
End If

If G9 = "拆一端" Then
    If G2 = "水冷" Then
        H9 = Round(300 * G4 * (1 + G7 / 1050), 0)
    Else
        H9 = 0
    End If
ElseIf G9 = "拆两端" Then
    If G2 = "水冷" Then
        H9 = Round(300 * G4 * (1 + G7 / 1050) * 1.1, 0)
    Else
        H9 = 0
    End If
Else
    H9 = 0
End If

If I9 = "拆一端" Then
    If I2 = "水冷" Then
        J9 = Round(300 * I4 * (1 + I7 / 1050), 0)
    Else
        J9 = 0
    End If
ElseIf I9 = "拆两端" Then
    If I2 = "水冷" Then
        J9 = Round(300 * I4 * (1 + I7 / 1050) * 1.1, 0)
    Else
        J9 = 0
    End If
Else
    J9 = 0
End If

If C10 = "拆一端" Then
    If C2 = "水冷" Then
        D10 = Round(300 * C4 * (1 + C7 / 1050), 0)
    Else
        D10 = 0
    End If
ElseIf C10 = "拆两端" Then
    If C2 = "水冷" Then
        D10 = Round(300 * C4 * (1 + C7 / 1050) * 1.1, 0)
    Else
        D10 = 0
    End If
Else
    D10 = 0
End If

If E10 = "拆一端" Then
    If E2 = "水冷" Then
        F10 = Round(300 * E4 * (1 + E7 / 1050), 0)
    Else
        F10 = 0
    End If
ElseIf E10 = "拆两端" Then
    If E2 = "水冷" Then
        F10 = Round(300 * E4 * (1 + E7 / 1050) * 1.1, 0)
    Else
        F10 = 0
    End If
Else
    F10 = 0
End If
If G10 = "拆一端" Then
    If G2 = "水冷" Then
        H10 = Round(300 * G4 * (1 + G7 / 1050), 0)
    Else
        H10 = 0
    End If
ElseIf G10 = "拆两端" Then
    If G2 = "水冷" Then
        H10 = Round(300 * G4 * (1 + G7 / 1050) * 1.1, 0)
    Else
        H10 = 0
    End If
Else
    H10 = 0
End If
If I10 = "拆一端" Then
    If I2 = "水冷" Then
        J10 = Round(300 * I4 * (1 + I7 / 1050), 0)
    Else
        J10 = 0
    End If
ElseIf I10 = "拆两端" Then
    If I2 = "水冷" Then
        J10 = Round(300 * I4 * (1 + I7 / 1050) * 1.1, 0)
    Else
        J10 = 0
    End If
Else
    J10 = 0
End If

'd11=IF(C2="水冷",0,IF(C11="需要",400*C4*(1+C7/350),0))
If C2 = "水冷" Then
    D11 = 0
Else
    If C11 = "需要" Then
        D11 = Round(400 * Val(C4) * (1 + Val(C7) / 350), 0)
    Else
        D11 = 0
    End If
End If
If E2 = "水冷" Then
    F11 = 0
Else
    If E11 = "需要" Then
        F11 = Round(400 * Val(E4) * (1 + Val(E7) / 350), 0)
    Else
        F11 = 0
    End If
End If
If G2 = "水冷" Then
    H11 = 0
Else
    If G11 = "需要" Then
        H11 = Round(400 * Val(G4) * (1 + Val(G7) / 350), 0)
    Else
        H11 = 0
    End If
End If
If I2 = "水冷" Then
    J11 = 0
Else
    If I11 = "需要" Then
        J11 = Round(400 * Val(I4) * (1 + Val(I7) / 350), 0)
    Else
        J11 = 0
    End If
End If
'd12=IF(C2="水冷",0,IF(C12="需要",100*C4*(1+C7/350),0))
If C2 = "水冷" Then
    D12 = 0
Else
    If C12 = "需要" Then
        D12 = Round(100 * Val(C4) * (1 + Val(C7) / 350), 0)
    Else
        D12 = 0
    End If
End If
If E2 = "水冷" Then
    F12 = 0
Else
    If E12 = "需要" Then
        F12 = Round(100 * Val(E4) * (1 + Val(E7) / 350), 0)
    Else
        F12 = 0
    End If
End If
If G2 = "水冷" Then
    H12 = 0
Else
    If G12 = "需要" Then
        H12 = Round(100 * Val(G4) * (1 + Val(G7) / 350), 0)
    Else
        H12 = 0
    End If
End If
If I2 = "水冷" Then
    J12 = 0
Else
    If I12 = "需要" Then
        J12 = Round(100 * Val(I4) * (1 + Val(I7) / 350), 0)
    Else
        J12 = 0
    End If
End If
If C13 = "需要" Then
    D13 = Round(400 * (1 + C7 / 1050), 0)
Else
    D13 = 0
End If
If E13 = "需要" Then
    F13 = Round(400 * (1 + E7 / 1050), 0)
Else
    F13 = 0
End If
If G13 = "需要" Then
    H13 = Round(400 * (1 + G7 / 1050), 0)
Else
    H13 = 0
End If
If I13 = "需要" Then
    J13 = Round(400 * (1 + I7 / 1050), 0)
Else
    J13 = 0
End If
If C14 = "需要" Then
    D14 = 600 * C4
Else
    D14 = 0
End If
If E14 = "需要" Then
    F14 = 600 * E4
Else
    F14 = 0
End If
If G14 = "需要" Then
    H14 = 600 * G4
Else
    H14 = 0
End If
If I14 = "需要" Then
    J14 = 600 * I4
Else
    J14 = 0
End If
If C15 = "需要" Then
    D15 = 600 * C4
Else
    D15 = 0
End If
If E15 = "需要" Then
    F15 = 600 * E4
Else
    F15 = 0
End If
If G15 = "需要" Then
    H15 = 600 * G4
Else
    H15 = 0
End If
If I15 = "需要" Then
    J15 = 600 * I4
Else
    J15 = 0
End If
If C16 = "需要" Then
    D16 = 400 * C4
Else
    D16 = 0
End If
If E16 = "需要" Then
    F16 = 400 * E4
Else
    F16 = 0
End If
If G16 = "需要" Then
    H16 = 400 * G4
Else
    H16 = 0
End If
If I16 = "需要" Then
    J16 = 400 * I4
Else
    J16 = 0
End If
If C17 = "需要" Then
    D17 = 400 * C4
Else
    D17 = 0
End If
If E17 = "需要" Then
    F17 = 400 * E4
Else
    F17 = 0
End If
If G17 = "需要" Then
    H17 = 400 * G4
Else
    H17 = 0
End If
If I17 = "需要" Then
    J17 = 400 * I4
Else
    J17 = 0
End If
'=1200*(C4+E4+G4+I4)*(1+((C7*C4+E7*E4+G7*G4+I7*I4)/(C4+E4+G4+I4)-1400)/2100)
D18 = Round(1200 * (Val(C4) + Val(E4) + Val(G4) + Val(I4)) * (1 + ((Val(C7) * Val(C4) + Val(E7) * Val(E4) + Val(G7) * Val(G4) + Val(I7) * Val(I4)) / (Val(C4) + Val(E4) + Val(G4) + Val(I4)) - 1400) / 2100), 0)
D19 = 350 * (Val(C19) + Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1)
'd21=600*(3+(C4+E4+G4+I4-1)*2)*(1+((C7*C4+E7*E4+G7*G4+I7*I4)/(C4+E4+G4+I4)-700)/2100))
If C2 = "热泵" Then
    D20 = Round(600 * (3 + (Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1) * 2) * (1 + ((Val(C7) * Val(C4) + Val(E7) * Val(E4) + Val(G7) * Val(G4) + Val(I7) * Val(I4)) / (Val(C4) + Val(E4) + Val(G4) + Val(I4)) - 700) / 1050) * 2, 0)
Else
    D20 = Round(600 * (3 + (Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1) * 2) * (1 + ((Val(C7) * Val(C4) + Val(E7) * Val(E4) + Val(G7) * Val(G4) + Val(I7) * Val(I4)) / (Val(C4) + Val(E4) + Val(G4) + Val(I4)) - 700) / 2100), 0)
End If

D21 = Round(((Val(D8) + Val(D9) + Val(D10) + Val(D11) + Val(D12) + Val(D13) + Val(D14) + Val(D15) + Val(D16) + Val(D17)) + _
        (Val(F8) + Val(F9) + Val(F10) + Val(F11) + Val(F12) + Val(F13) + Val(F14) + Val(F15) + Val(F16) + Val(F17)) + _
        (Val(H8) + Val(H9) + Val(H10) + Val(H11) + Val(H12) + Val(H13) + Val(H14) + Val(H15) + Val(H16) + Val(H17)) + _
        (Val(J8) + Val(J9) + Val(J10) + Val(J11) + Val(J12) + Val(J13) + Val(J14) + Val(J15) + Val(J16) + Val(J17)) + Val(D18) + Val(D19) + Val(D20)) * 1.2, 0)
If D22.Text = "全包" Then
    D21 = Round((D21 * 1.5), 0)
End If

'差旅
N3 = 30 + 2 * Val(N2)
'n4=N3*(1+(C4+E4+G4+I4-1)/2)
N4 = Val(N3) * (1 + (Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1) / 2)
'N5 = Val(N3) * (Val(C20) + Val(C5) + Val(E5) + Val(G5) + Val(I5) - 1)
N5 = Val(N3) * (Val(C19) + Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1)
'=N3*(C19+C4+E4+G4+I4-1)
'=N3*(3+(C4+E4+G4+I4-1)*2)
N6 = Val(N3) * (3 + (Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1) * 2)
O4 = Val(O3) * (1 + (Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1) / 2)
O5 = 0: P5 = 0
O6 = Val(O3) * (1 + (Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1) / 2)
P4 = Val(P3) * (2 + (Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1) / 2)
P6 = Val(P3) * (1 + (Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1) / 2)
Q4 = Val(Q3) * (2 + (Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1) / 2)
Q5 = Val(Q3) * (Val(C19) + Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1)
Q6 = Val(Q3) * (3 + (Val(C4) + Val(E4) + Val(G4) + Val(I4) - 1) * 2)
N7 = Val(N4) + Val(N5) + Val(N6)
O7 = Val(O4) + Val(O5) + Val(O6)
P7 = Val(P4) + Val(P5) + Val(P6)
Q7 = Val(Q4) + Val(Q5) + Val(Q6)
If Val(N2.Text) > 260 Then
    N8 = Val(N7) + Val(Q7)
ElseIf Val(N2) > 40 Then
    N8 = Val(N7) + Val(P7)
ElseIf Val(N2) > 20 Then
    N8 = Val(N7) + Val(O7)
Else
    N8 = Val(N7)
End If
'n8=N7+IF(N2>260,Q7,IF(N2>40,P7,IF(N2>20,O7,0)))
End Sub

