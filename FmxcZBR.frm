VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcZBR 
   BackColor       =   &H00C0FFC0&
   Caption         =   "成本追加一览"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   375
      Left            =   870
      TabIndex        =   3
      Top             =   2700
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   661
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "添加"
      Height          =   585
      Left            =   3390
      Picture         =   "FmxcZBR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2490
      Width           =   675
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "关闭"
      Height          =   585
      Left            =   4080
      Picture         =   "FmxcZBR.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2490
      Width           =   585
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgZBr 
      Height          =   2445
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   4313
      _Version        =   393216
      BackColor       =   12648384
      BackColorFixed  =   12648384
      BackColorBkg    =   12648447
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "FmxcZBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub cmdSave_Click()
Dim YZE As Single
Dim XZE As Single
'''''If FmxcNew.Visible = True Then '检验是否超出预估成本总额
'''''    YZE = Val(FmxcNew.txtQb.Text)
'''''    dtgN.Col = 2
'''''    dtgN
'''''End If
If mod1.DName <> "朱婷婷" Then
    Exit Sub
End If
Call fmxcZJ.Qing
If FMXC.Visible = True Then
'引用合同数据
fmxcZJ.lblKhmc.Caption = FMXC.txtKhmc.Text
fmxcZJ.lblGLBH.Caption = FMXC.txtHtbh.Text
fmxcZJ.lblGLBH.ToolTipText = FMXC.lblMHid.Caption
fmxcZJ.lblZbh.Caption = FMXC.txtZbh.Text
fmxcZJ.lblXz.Caption = FMXC.lblHtxz.Caption
fmxcZJ.lblZE.Caption = FMXC.txtHtze.Text
ElseIf FmxcNew.Visible = True Then
    fmxcZJ.lblKhmc.Caption = FmxcNew.txtKhmc.Text
    fmxcZJ.lblGLBH.Caption = FmxcNew.txtHtbh.Text
    fmxcZJ.lblGLBH.ToolTipText = FmxcNew.lblHid.Caption
    fmxcZJ.lblZbh.Caption = FmxcNew.txtZbh.Text
    fmxcZJ.lblXz.Caption = FmxcNew.txtHtxz.Text
    fmxcZJ.lblZE.Caption = FmxcNew.txtHtze.Text
End If
fmxcZJ.cmdSave.Enabled = True

fmxcZJ.lblYwy.Caption = mod1.DName
Call fmxcZJ.dtgPFF
fmxcZJ.Show
fmxcZJ.ZOrder 0
Me.Visible = False
fmxcZJ.frmGui.Visible = True
fmxcZJ.frmGui.Enabled = True
fmxcZJ.comGui.Enabled = True
fmxcZJ.optF.Enabled = True
fmxcZJ.optQ.Enabled = True
End Sub

Private Sub dtgZBr_DblClick()

dtgN.Row = dtgZBr.Row
dtgN.Col = 3
Call fmxcZJ.Bound(Val(dtgN.Text))
fmxcZJ.Show
fmxcZJ.ZOrder 0
Me.Visible = False
End Sub


Private Sub Form_Load()
Me.Height = 3600
Me.Width = 4800
Call dtgFF

End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Cancel = True
End Sub

Public Sub dtgFF() '表格初始,格式
dtgZBr.Clear
dtgZBr.Cols = 4: dtgZBr.Rows = 200
dtgZBr.ColWidth(3) = 0
dtgZBr.ColWidth(0) = 2280
dtgZBr.Row = 0
dtgZBr.Col = 0: dtgZBr.Text = "编号": dtgZBr.CellFontBold = True
dtgZBr.Col = 1: dtgZBr.Text = "费用归属": dtgZBr.CellFontBold = True
dtgZBr.Col = 2: dtgZBr.Text = "成本金额": dtgZBr.CellFontBold = True

dtgN.Clear
dtgN.Cols = 4: dtgN.Rows = 200
dtgN.ColWidth(3) = 0
dtgN.ColWidth(0) = 2280
dtgN.Row = 0
dtgN.Col = 0: dtgN.Text = "编号": dtgN.CellFontBold = True
dtgN.Col = 1: dtgN.Text = "费用归属": dtgN.CellFontBold = True
dtgN.Col = 2: dtgN.Text = "成本金额": dtgN.CellFontBold = True
End Sub
