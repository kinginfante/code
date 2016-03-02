VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGxbjX 
   BackColor       =   &H00C0FFC0&
   Caption         =   "销售询价"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   10215
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdGB 
      Caption         =   "关闭"
      Height          =   345
      Left            =   9120
      TabIndex        =   9
      Top             =   6030
      Width           =   1065
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "frmGxbjX.frx":0000
      Left            =   1500
      List            =   "frmGxbjX.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   6030
      Width           =   2115
   End
   Begin VB.TextBox txtZ 
      Height          =   285
      Left            =   4230
      TabIndex        =   7
      Top             =   6030
      Width           =   1905
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "查询"
      Height          =   315
      Left            =   6360
      TabIndex        =   6
      Top             =   6060
      Width           =   915
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "全方式查询"
      Height          =   315
      Left            =   7350
      TabIndex        =   5
      ToolTipText     =   $"frmGxbjX.frx":0038
      Top             =   6060
      Width           =   1725
   End
   Begin VB.TextBox txtSl 
      Height          =   270
      Left            =   8430
      TabIndex        =   3
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton cmdDao 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导入"
      Height          =   285
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   1035
   End
   Begin VB.Timer timQuit 
      Interval        =   1000
      Left            =   630
      Top             =   90
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtJzxh 
      Height          =   270
      Left            =   5490
      TabIndex        =   1
      Top             =   5640
      Width           =   2205
   End
   Begin VB.ComboBox comJzPb 
      Height          =   300
      ItemData        =   "frmGxbjX.frx":006A
      Left            =   960
      List            =   "frmGxbjX.frx":0080
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5640
      Width           =   1875
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgHP 
      Height          =   4515
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   7964
      _Version        =   393216
      BackColor       =   12648447
      BackColorFixed  =   12648384
      BackColorBkg    =   12648384
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "查询方式"
      Height          =   255
      Left            =   300
      TabIndex        =   15
      Top             =   6060
      Width           =   825
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "值"
      Height          =   195
      Left            =   3720
      TabIndex        =   14
      Top             =   6090
      Width           =   345
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "数量"
      Height          =   225
      Left            =   7890
      TabIndex        =   13
      Top             =   5670
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "机组品牌"
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   5700
      Width           =   795
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "机组型号"
      Height          =   225
      Left            =   4590
      TabIndex        =   11
      Top             =   5700
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10200
      Y1              =   6000
      Y2              =   6000
   End
End
Attribute VB_Name = "frmGxbjX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dtgHP.Rows = 50
dtgHP.Cols = 7
dtgHP.ColWidth(0) = 300: dtgHP.ColWidth(6) = 0

dtgHP.ColWidth(1) = 1380
dtgHP.ColWidth(2) = 2970
dtgHP.ColWidth(3) = 2355
dtgHP.ColWidth(4) = 1800
'dtgHp.ColWidth(5) = 1380
Me.Height = 6945: Me.Width = 10350
End Sub
