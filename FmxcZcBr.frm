VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcZcBr 
   BackColor       =   &H00C0FFC0&
   Caption         =   "付款查询"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15060
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   15060
   Begin VB.ComboBox comXZ1 
      Height          =   300
      ItemData        =   "FmxcZcBr.frx":0000
      Left            =   4080
      List            =   "FmxcZcBr.frx":0010
      TabIndex        =   12
      Text            =   "全部"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H008080FF&
      Caption         =   "超　期"
      Height          =   315
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdSerach 
      Caption         =   "查询"
      Height          =   315
      Left            =   5520
      TabIndex        =   10
      Top             =   4680
      Width           =   945
   End
   Begin VB.CommandButton cmdW 
      BackColor       =   &H00FFC0C0&
      Caption         =   "未到帐"
      Height          =   315
      Left            =   5445
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5190
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSComCtl2.DTPicker dtpL 
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   126025729
      CurrentDate     =   41725
   End
   Begin MSComCtl2.DTPicker dtpF 
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Format          =   116260865
      CurrentDate     =   41725
   End
   Begin VB.TextBox txtZ1 
      Height          =   270
      Left            =   2640
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   375
      Left            =   10440
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox comTj 
      Height          =   300
      ItemData        =   "FmxcZcBr.frx":002E
      Left            =   960
      List            =   "FmxcZcBr.frx":0047
      TabIndex        =   3
      Text            =   "供应商名称"
      Top             =   4680
      Width           =   1605
   End
   Begin VB.TextBox txtZ 
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2640
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制"
      Height          =   315
      Left            =   7950
      TabIndex        =   1
      ToolTipText     =   "点击后，打开Excel，可进行粘贴"
      Top             =   5190
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   7646
      _Version        =   393216
      BackColor       =   16777152
      FixedCols       =   0
      BackColorFixed  =   15728356
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "查询条件"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   4710
      Width           =   1005
   End
End
Attribute VB_Name = "FmxcZcBr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ETT As String
Dim ZCid As Long

Private Sub cmdC_Click()
    Me.ETT = "select gymc,yhh,htbh,fkrq,fkje,cgy,fph,qrq,qje,fpdrq,fpwyy,zcid from zcbView where wcf=0 and fpdrq<'" & mod1.DQda & "'"
    
    Call Me.Bound(Me.ETT)
End Sub

Private Sub cmdCopy_Click()
If Not (mod1.DName = "顾珣" Or mod1.DName = "乔继敏" Or mod1.DName = "张文琴" Or mod1.DName = "于晓静" Or mod1.DName = "马晓聪" Or mod1.DName = "陈文超" Or mod1.DName = "张萍") Then
    Exit Sub
End If
dtgBr.FixedCols = 0
dtgBr.FixedRows = 0
dtgBr.Row = 0
dtgBr.Col = 0
dtgBr.ColSel = 13
dtgBr.RowSel = dtgBr.Rows - 1
Clipboard.Clear
Clipboard.SetText dtgBr.Clip
dtgBr.FixedCols = 1
dtgBr.FixedRows = 1
End Sub

Private Sub cmdSerach_Click()

Select Case comTj.Text
Case "供应商名称"
    Me.ETT = "select company,gymc,yhh,htbh,fkrq,fkje,cgy,qrq,fph,qje,fpdrq,fpwyy,zcid,wcf from zcbView where gymc like '%" & txtZ.Text & "%'"
Case "合同号"
    Me.ETT = "select company,gymc,yhh,htbh,fkrq,fkje,cgy,qrq,fph,qje,fpdrq,fpwyy,zcid,wcf from zcbView where htbh like '%" & txtZ.Text & "%'"
Case "付款日期"
    Me.ETT = "select company,gymc,yhh,htbh,fkrq,fkje,cgy,qrq,fph,qje,fpdrq,fpwyy,zcid,wcf from zcbView where fkrq >='" & _
    DateSerial(Year(txtZ.Text), Month(txtZ.Text), Day(txtZ.Text)) & "' and fkrq < '" & _
    DateSerial(Year(txtZ1.Text), Month(txtZ1.Text), Day(txtZ1.Text) + 1) & "'"
Case "付款金额"
    Me.ETT = "select company,gymc,yhh,htbh,fkrq,fkje,cgy,qrq,fph,qje,fpdrq,fpwyy,zcid,wcf from zcbView where fkje =" & txtZ.Text
Case "采购员"
    Me.ETT = "select company,gymc,yhh,htbh,fkrq,fkje,cgy,qrq,fph,qje,fpdrq,fpwyy,zcid,wcf from zcbView where cgy like '%" & txtZ.Text & "%'"
Case "发票号"
    Me.ETT = "select company,gymc,yhh,htbh,fkrq,fkje,cgy,qrq,fph,qje,fpdrq,fpwyy,zcid,wcf from zcbView where fph like '%" & txtZ.Text & "%'"
Case "发票暂支到期日"
    Me.ETT = "select company,gymc,yhh,htbh,fkrq,fkje,cgy,qrq,fph,qje,fpdrq,fpwyy,zcid,wcf from zcbView where fkdrq >='" & _
    DateSerial(Year(txtZ.Text), Month(txtZ.Text), Day(txtZ.Text)) & "' and fkdrq < '" & _
    DateSerial(Year(txtZ1.Text), Month(txtZ1.Text), Day(txtZ1.Text) + 1) & "'"
End Select

If txtZ.Text = "" Then
    Select Case comXZ1.Text
    Case "完成"
        Me.ETT = "select company,gymc,yhh,htbh,fkrq,fkje,cgy,qrq,fph,qje,fpdrq,fpwyy,zcid,wcf from zcbView where wcf=1 order by zcid,qid"
    Case "未完成"
        Me.ETT = "select company,gymc,yhh,htbh,fkrq,fkje,cgy,qrq,fph,qje,fpdrq,fpwyy,zcid,wcf from zcbView where wcf=0 order by zcid,qid"
    Case "超期"
        Me.ETT = "select company,gymc,yhh,htbh,fkrq,fkje,cgy,qrq,fph,qje,fpdrq,fpwyy,zcid,wcf from zcbView where fpdrq<'" & mod1.DQda & "' order by zcid,qid"
    Case "全部"
        Me.ETT = "select company,gymc,yhh,htbh,fkrq,fkje,cgy,qrq,fph,qje,fpdrq,fpwyy,zcid,wcf from zcbView  order by zcid,qid"
    End Select
Else

    Select Case comXZ1.Text
    Case "完成"
        Me.ETT = Me.ETT & " and wcf=1 order by zcid,qid"
    Case "未完成"
        Me.ETT = Me.ETT & " and wcf=0 order by zcid,qid"
    Case "超期"
        Me.ETT = Me.ETT & " and fpdrq<'" & mod1.DQda & "' order by zcid,qid"
    Case "全部"
        Me.ETT = Me.ETT & " order by zcid,qid"
    End Select
End If
Call Me.Bound(Me.ETT)
End Sub

Private Sub cmdW_Click()
    Me.ETT = "select gymc,yhh,htbh,fkrq,fkje,cgy,qrq,fph,qje,fpdrq,fpwyy,zcid from zcbView where wcf=0"
    Call Me.Bound(Me.ETT)
End Sub

Private Sub comTj_Click()
txtZ.Visible = True
txtZ1.Visible = False
dtpL.Visible = False
dtpF.Visible = False
Select Case comTj.Text
Case "供应商名称"

Case "合同号"

Case "付款日期"
    txtZ.Text = DateSerial(Year(dtpF.Value), Month(dtpF.Value), Day(dtpF.Value))
    txtZ1.Text = DateSerial(Year(dtpL.Value), Month(dtpL.Value), Day(dtpL.Value))
    txtZ.Visible = False: txtZ1.Visible = False
    dtpF.Visible = True: dtpL.Visible = True
Case "付款金额"

Case "采购员"

Case "发票号"

Case "发票暂支到期日"
    txtZ.Text = DateSerial(Year(dtpF.Value), Month(dtpF.Value), Day(dtpF.Value))
    txtZ1.Text = DateSerial(Year(dtpL.Value), Month(dtpL.Value), Day(dtpL.Value))
    txtZ.Visible = False: txtZ1.Visible = False
    dtpF.Visible = True: dtpL.Visible = True
End Select
End Sub

Private Sub dtgBr_Click()
On Error Resume Next
dtgN.Row = dtgBr.Row
dtgN.Col = 12
ZCid = Val(dtgN.Text)
If ZCid = 0 Then Exit Sub
Call fmxcZC.Bound(ZCid)
End Sub

Private Sub dtpF_CloseUp()
    txtZ.Text = DateSerial(Year(dtpF.Value), Month(dtpF.Value), Day(dtpF.Value))
    txtZ.Visible = True
    dtpF.Visible = False
End Sub


Private Sub dtpL_CloseUp()
    txtZ1.Text = DateSerial(Year(dtpL.Value), Month(dtpL.Value), Day(dtpL.Value))
    txtZ1.Visible = True
    dtpL.Visible = False
End Sub


Private Sub Form_Click()
    dtpF.Visible = True: dtpL.Visible = True
End Sub

Private Sub Form_Load()
Me.Height = mod1.FHeight - 3000
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 3660
Call Me.dtgbrFF
dtpF.Value = mod1.DQda
dtpL.Value = mod1.DQda

End Sub

Public Sub dtgbrFF()
dtgBr.Clear
dtgBr.Cols = 14
dtgBr.Rows = 100
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "公司名称": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "供应商名称": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "银行流水号": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "合同号": dtgBr.CellFontBold = True
dtgBr.Col = 4: dtgBr.Text = "付款日期": dtgBr.CellFontBold = True
dtgBr.Col = 5: dtgBr.Text = "付款金额": dtgBr.CellFontBold = True
dtgBr.Col = 6: dtgBr.Text = "采购员": dtgBr.CellFontBold = True
dtgBr.Col = 7: dtgBr.Text = "清暂支日期": dtgBr.CellFontBold = True
dtgBr.Col = 8: dtgBr.Text = "发票号": dtgBr.CellFontBold = True
dtgBr.Col = 9: dtgBr.Text = "清暂支金额": dtgBr.CellFontBold = True
dtgBr.Col = 10: dtgBr.Text = "发票暂支到期日": dtgBr.CellFontBold = True
dtgBr.Col = 11: dtgBr.Text = "发票未到原因": dtgBr.CellFontBold = True
dtgBr.Col = 13: dtgBr.Text = "完成": dtgBr.CellFontBold = True
dtgBr.Col = 12: dtgBr.Text = "ZCid"
dtgBr.ColWidth(0) = 2730
dtgBr.ColWidth(1) = 1230: dtgBr.ColWidth(2) = 1275: dtgBr.ColWidth(3) = 990
dtgBr.ColWidth(7) = 1440: dtgBr.ColWidth(8) = 1110: dtgBr.ColWidth(9) = 1125: dtgBr.ColWidth(10) = 1575
dtgBr.ColWidth(11) = 1320: dtgBr.ColWidth(12) = 675: dtgBr.ColWidth(12) = 0
dtgN.Clear
dtgN.Cols = 14
dtgN.Rows = 100



End Sub

Public Sub Bound(tt As String)
Dim oo As Long
Dim ii As Integer
Dim Ra
Dim La As Long
If tt = "" Then Exit Sub
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgBr.Visible = False

Call Me.dtgbrFF
dtgBr.Rows = La + 50
dtgN.Rows = La + 50
For oo = 1 To La
    dtgBr.Row = oo: dtgN.Row = oo
    For ii = 0 To 13
        dtgBr.Col = ii: dtgN.Col = ii
        dtgBr.Text = Ra(ii, oo - 1)
        dtgN.Text = Ra(ii, oo - 1)
        If ii = 13 Then
            If Ra(ii, oo - 1) = False Then
            dtgBr.Text = "未完成"
            Else
            dtgBr.Text = "完成"
            End If
        End If
    Next
Next
dtgBr.Visible = True
End Sub

Private Sub txtZ_Click()
Select Case Me.comTj.Text
Case "供应商名称"

Case "合同号"

Case "付款日期"
    txtZ.Text = DateSerial(Year(dtpF.Value), Month(dtpF.Value), Day(dtpF.Value))
    txtZ1.Text = DateSerial(Year(dtpL.Value), Month(dtpL.Value), Day(dtpL.Value))
    txtZ.Visible = False
    dtpF.Visible = True
Case "付款金额"

Case "采购员"

Case "发票号"

Case "发票暂支到期日"
    txtZ.Text = DateSerial(Year(dtpF.Value), Month(dtpF.Value), Day(dtpF.Value))
    txtZ1.Text = DateSerial(Year(dtpL.Value), Month(dtpL.Value), Day(dtpL.Value))
    txtZ1.Visible = False
    dtpL.Visible = True
End Select
End Sub
