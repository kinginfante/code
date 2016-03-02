VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form fmxcXjBr 
   Caption         =   "询价单列表"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   915
      Left            =   2610
      TabIndex        =   1
      Top             =   450
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   1614
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   5477
      _Version        =   393216
      BackColor       =   16777152
      FixedCols       =   0
      BackColorFixed  =   12648384
      BackColorBkg    =   16777152
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblHid 
      Caption         =   "Label1"
      Height          =   315
      Left            =   3030
      TabIndex        =   2
      Top             =   1590
      Width           =   975
   End
End
Attribute VB_Name = "fmxcXjBr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub dtgFF()
dtgBr.Clear: dtgN.Clear
 dtgBr.Cols = 4
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "报价有效期": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "基准价": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "编号": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "bid": dtgBr.CellFontBold = True
dtgBr.ColWidth(3) = 0
 dtgN.Cols = 4
dtgBr.Rows = 50: dtgN.Rows = 50
lblHid.Caption = 0
dtgBr.ColWidth(0) = 2280
End Sub

Private Sub dtgBr_DblClick()
dtgN.Row = dtgBr.Row
dtgN.Col = 3
If Val(dtgN.Text) = 0 Then Exit Sub
Call FmxcXJ.Bound(Val(dtgN.Text))
FmxcXJ.Show
FmxcXJ.ZOrder
FmxcXJ.cmdHT.ToolTipText = lblHid.Caption
FmxcXJ.cmdDht.Visible = True
MsgBox ("如果您点了'导入合同',此询价单将与合同评审单（尾号" & lblHid.Caption & "）关联！")

End Sub

Private Sub Form_Load()
Me.Height = 3600
Me.Width = 4800
End Sub
