VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmHtZX 
   BackColor       =   &H00C0FFC0&
   Caption         =   "合同执行列表"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14520
      Picture         =   "frmHtZX.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8580
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8355
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   14737
      _Version        =   393216
      BackColor       =   12648384
      Rows            =   10
      Cols            =   8
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      BackColorUnpopulated=   8454016
      GridColorUnpopulated=   8454016
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
End
Attribute VB_Name = "frmHtZX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fwid As Long

Private Sub cmdBack_Click()
frmZu.Enabled = True
Me.Visible = False
End Sub

Private Sub dtgBr_DblClick()
Dim Zid As Long

Zid = Val(dtgBr.Text)
        Call frmHtz1.Qing
        Call frmHtz1.Bound(0, Zid)
        Call frmHtz1.dtgFF
        
        frmHtz1.Show
End Sub

Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
dtgBr.ColWidth(1) = 2025
dtgBr.ColWidth(2) = 2700
dtgBr.ColWidth(4) = 4245
dtgBr.ColWidth(5) = 1755

dtgBr.Rows = 30

dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "单号": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "合同编号": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "项目名称": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "类型": dtgBr.CellFontBold = True
dtgBr.Col = 4: dtgBr.Text = "内容": dtgBr.CellFontBold = True
dtgBr.Col = 5: dtgBr.Text = "发布日期": dtgBr.CellFontBold = True
dtgBr.Col = 6: dtgBr.Text = "付款": dtgBr.CellFontBold = True
dtgBr.Col = 7: dtgBr.Text = "执行状态": dtgBr.CellFontBold = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
frmZu.Enabled = True
Cancel = True

End Sub

Public Sub Bound(RA, La As Long)
Dim oo As Long
Dim ii As Long
On Error Resume Next
dtgBr.Visible = False
dtgBr.Rows = La + 30
For oo = 1 To La
    dtgBr.Row = oo
    For ii = 0 To 7
        dtgBr.Col = ii
        dtgBr.Text = RA(ii, oo - 1)
    Next
Next
dtgBr.Visible = True
End Sub

