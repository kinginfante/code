VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGxbjSD 
   BackColor       =   &H00C0FFC0&
   Caption         =   "速达货品库存"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   10215
   Begin VB.ComboBox comJzPb 
      Height          =   300
      ItemData        =   "frmGxbjSD.frx":0000
      Left            =   2400
      List            =   "frmGxbjSD.frx":0016
      TabIndex        =   15
      Top             =   5640
      Width           =   1875
   End
   Begin VB.TextBox txtJzxh 
      Height          =   270
      Left            =   5490
      TabIndex        =   14
      Top             =   5640
      Width           =   2205
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timQuit 
      Interval        =   1000
      Left            =   630
      Top             =   90
   End
   Begin VB.CommandButton cmdDao 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导入"
      Height          =   285
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   1035
   End
   Begin VB.TextBox txtSl 
      Height          =   270
      Left            =   8430
      TabIndex        =   9
      Top             =   5640
      Width           =   615
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   315
      Left            =   30
      TabIndex        =   8
      Top             =   5610
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "全方式查询"
      Height          =   315
      Left            =   7350
      TabIndex        =   7
      ToolTipText     =   $"frmGxbjSD.frx":0046
      Top             =   6060
      Width           =   1725
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "查询"
      Height          =   315
      Left            =   6360
      TabIndex        =   6
      Top             =   6060
      Width           =   915
   End
   Begin VB.TextBox txtZ 
      Height          =   285
      Left            =   4230
      TabIndex        =   5
      Top             =   6030
      Width           =   1905
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "frmGxbjSD.frx":0078
      Left            =   1500
      List            =   "frmGxbjSD.frx":0085
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   6030
      Width           =   2115
   End
   Begin VB.CommandButton cmdGB 
      Caption         =   "关闭"
      Height          =   345
      Left            =   9120
      TabIndex        =   1
      Top             =   6030
      Width           =   1065
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgHP 
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   9869
      _Version        =   393216
      BackColor       =   12648447
      BackColorFixed  =   12648384
      BackColorBkg    =   12648384
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10200
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "机组型号"
      Height          =   225
      Left            =   4590
      TabIndex        =   13
      Top             =   5700
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "机组品牌"
      Height          =   225
      Left            =   1560
      TabIndex        =   12
      Top             =   5700
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "数量"
      Height          =   225
      Left            =   7890
      TabIndex        =   11
      Top             =   5670
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "值"
      Height          =   195
      Left            =   3720
      TabIndex        =   4
      Top             =   6090
      Width           =   345
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "查询方式"
      Height          =   255
      Left            =   300
      TabIndex        =   2
      Top             =   6060
      Width           =   825
   End
End
Attribute VB_Name = "frmGxbjSD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CT As String
Dim timZm As Integer '1导入


Private Sub cmdAll_Click()
If comJzPb.Text = "" Then
    MsgBox "请选择机组品牌!"
    Exit Sub
End If

    CT = "SELECT dbo.l_goods.code, dbo.l_goods.name, dbo.l_goods.specs, dbo.l_goodstype.name AS goodtypename, dbo.l_goodsunit.unitname,dbo.l_goods.goodsid" & _
        " FROM dbo.l_goods LEFT OUTER JOIN dbo.l_goodsunit ON dbo.l_goods.goodsid = dbo.l_goodsunit.goodsid LEFT OUTER JOIN dbo.l_goodstype ON dbo.l_goods.gdtypeid = dbo.l_goodstype.gdtypeid" & _
        " where dbo.l_goods.closed=0 and  (dbo.l_goods.name like '%" & txtZ.Text & "%' or dbo.l_goods.specs like '%" & txtZ.Text & "%') and (dbo.l_goodstype.name like '%" & comJzPb.Text & "%' or dbo.l_goodstype.name like '%通用%')"
Call dtgFF
Call CX(CT)
txtSL.Text = ""
End Sub

Private Sub cmdC_Click()
If comLx.Text = "" Then
    comLx.Text = "货品名称"
End If
If comJzPb.Text = "" And mod1.ZT <> "HBData" Then
    MsgBox "请选择机组品牌!"
    Exit Sub
End If
Select Case comLx.Text
Case "货品名称"
    CT = "SELECT dbo.l_goods.code, dbo.l_goods.name, dbo.l_goods.specs, dbo.l_goodstype.name AS goodtypename, dbo.l_goodsunit.unitname,dbo.l_goods.goodsid" & _
        " FROM dbo.l_goods LEFT OUTER JOIN dbo.l_goodsunit ON dbo.l_goods.goodsid = dbo.l_goodsunit.goodsid LEFT OUTER JOIN dbo.l_goodstype ON dbo.l_goods.gdtypeid = dbo.l_goodstype.gdtypeid" & _
        " where dbo.l_goods.closed=0 and dbo.l_goods.name like '%" & txtZ.Text & "%'  and (dbo.l_goodstype.name like '%" & comJzPb.Text & "%' or dbo.l_goodstype.name like '%通用%')"
        If mod1.ZT = "HBData" Then
            CT = "SELECT dbo.l_goods.code, dbo.l_goods.name, dbo.l_goods.specs, dbo.l_goodstype.name AS goodtypename, dbo.l_goodsunit.unitname,dbo.l_goods.goodsid" & _
                " FROM dbo.l_goods LEFT OUTER JOIN dbo.l_goodsunit ON dbo.l_goods.goodsid = dbo.l_goodsunit.goodsid LEFT OUTER JOIN dbo.l_goodstype ON dbo.l_goods.gdtypeid = dbo.l_goodstype.gdtypeid" & _
                " where dbo.l_goods.closed=0 and dbo.l_goods.name like '%" & txtZ.Text & "%' "
        End If
Case "规格"
    CT = "SELECT dbo.l_goods.code, dbo.l_goods.name, dbo.l_goods.specs, dbo.l_goodstype.name AS goodtypename, dbo.l_goodsunit.unitname,dbo.l_goods.goodsid" & _
        " FROM dbo.l_goods LEFT OUTER JOIN dbo.l_goodsunit ON dbo.l_goods.goodsid = dbo.l_goodsunit.goodsid LEFT OUTER JOIN dbo.l_goodstype ON dbo.l_goods.gdtypeid = dbo.l_goodstype.gdtypeid" & _
        " where dbo.l_goods.closed=0 and  dbo.l_goods.specs like '%" & txtZ.Text & "%' and (dbo.l_goodstype.name like '%" & comJzPb.Text & "%' or dbo.l_goodstype.name like '%通用%')"
        If mod1.ZT = "HBData" Then
            CT = "SELECT dbo.l_goods.code, dbo.l_goods.name, dbo.l_goods.specs, dbo.l_goodstype.name AS goodtypename, dbo.l_goodsunit.unitname,dbo.l_goods.goodsid" & _
                " FROM dbo.l_goods LEFT OUTER JOIN dbo.l_goodsunit ON dbo.l_goods.goodsid = dbo.l_goodsunit.goodsid LEFT OUTER JOIN dbo.l_goodstype ON dbo.l_goods.gdtypeid = dbo.l_goodstype.gdtypeid" & _
                " where dbo.l_goods.closed=0 and  dbo.l_goods.specs like '%" & txtZ.Text & "%' "
        End If
Case "货品类别"
    CT = "SELECT dbo.l_goods.code, dbo.l_goods.name, dbo.l_goods.specs, dbo.l_goodstype.name AS goodtypename, dbo.l_goodsunit.unitname,dbo.l_goods.goodsid" & _
        " FROM dbo.l_goods LEFT OUTER JOIN dbo.l_goodsunit ON dbo.l_goods.goodsid = dbo.l_goodsunit.goodsid LEFT OUTER JOIN dbo.l_goodstype ON dbo.l_goods.gdtypeid = dbo.l_goodstype.gdtypeid" & _
        " where dbo.l_goods.closed=0 and  dbo.l_goodstype.name like '%" & txtZ.Text & "%'"
Case "货品编码"
    CT = "SELECT dbo.l_goods.code, dbo.l_goods.name, dbo.l_goods.specs, dbo.l_goodstype.name AS goodtypename, dbo.l_goodsunit.unitname,dbo.l_goods.goodsid" & _
        " FROM dbo.l_goods LEFT OUTER JOIN dbo.l_goodsunit ON dbo.l_goods.goodsid = dbo.l_goodsunit.goodsid LEFT OUTER JOIN dbo.l_goodstype ON dbo.l_goods.gdtypeid = dbo.l_goodstype.gdtypeid" & _
        " where dbo.l_goods.closed=0 and  dbo.l_goods.code like '%" & txtZ.Text & "%'"
End Select
Call dtgFF
Call CX(CT)
txtSL.Text = ""
End Sub




Private Sub cmdDao_Click()
On Error Resume Next
Dim MC As String '名称
Dim Dw As String '单位
Dim GoodsCode As String '编码
Dim GG As String '规格
Dim GoodsId As String '速达系统ID
Dim hg As Long
'If comLx.Text = "" Then Exit Sub
If frmGXBj.Visible = True Then
    If frmGXBj.lblZl.Caption = "配件" Then
        frmGXBj.comLx.Text = "零配件"
    Else
        frmGXBj.comLx.Text = "产品"
    End If
    
    If comJzPb.Text = "" Then
        MsgBox "请确认机组品牌!"
        comJzPb.SetFocus
        Exit Sub
    End If
    If txtJzxh.Text = "" Then
        MsgBox "请确认机组型号!"
        txtJzxh.SetFocus
        Exit Sub
    End If
    If Val(txtSL.Text) = 0 Then
        MsgBox "请确认数量!"
        txtSL.SetFocus
        Exit Sub
    End If
    If frmGXBj.txtDrq.Text = "" Then
        frmGXBj.txtDrq.Text = mod1.DQda
    End If
    If frmGXBj.txtBrq.Text = "" Then
        frmGXBj.txtBrq.Text = mod1.DQda
    End If
    
    
    dtgHP.Col = 1: GoodsCode = dtgHP.Text
    dtgHP.Col = 2: MC = Left(dtgHP.Text, 50)
    dtgHP.Col = 3: GG = Left(dtgHP.Text, 50)
    dtgHP.Col = 4: Lb = dtgHP.Text
    dtgHP.Col = 5: Dw = dtgHP.Text
    dtgHP.Col = 6: GoodsId = dtgHP.Text
    If Val(GoodsId) = 0 Then Exit Sub
                                       '新版本速达
        timZm = 1
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "询价单"
        mod1.cmd.Parameters("@NBLX") = "配件添加"
        mod1.cmd.Parameters("@bh") = frmGXBj.lblHtbh.Caption
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = frmGXBj.lblBid.Caption
        mod1.cmd.Parameters("@mt2") = frmGXBj.lblZl.Caption
        mod1.cmd.Parameters("@mt3") = comJzPb.Text  '机组品牌
        mod1.cmd.Parameters("@mt4") = txtJzxh.Text  '机组型号
        mod1.cmd.Parameters("@mt5") = Dw '压缩机型号,单位
        mod1.cmd.Parameters("@mt6") = "" '出厂编号
        mod1.cmd.Parameters("@mt7") = GoodsId '机组序列号
        mod1.cmd.Parameters("@mt8") = MC '零件名称
        mod1.cmd.Parameters("@mt9") = GoodsCode '零件规格号
        mod1.cmd.Parameters("@mt10") = GG '品牌及产地
        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1") = Val(txtSL.Text) '数量
        mod1.cmd.Parameters("@mb1") = 0
        If mod1.Bm = "配送中心" Then
        mod1.cmd.Parameters("@mb5") = 1
        Else
        mod1.cmd.Parameters("@mb5") = 0
        End If
        mod1.cmd.Parameters("@md1") = Null
        Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "成功" Then
            MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
            Exit Sub
        Else '提交成功,等待系统中心处理数据
            cmdAdd.Enabled = False
            cmdJG.Enabled = False
            Me.Enabled = False
            frmWaitA.Visible = True
            frmWaitA.Timer2.Enabled = False
    
            frmWaitA.ZOrder 0
            frmWaitA.Timer2.Enabled = True
            timWait.Enabled = True
        End If
    
    
        Set mod1.cmd = Nothing
ElseIf FmxcXJ.Visible = True Then
    If comJzPb.Text = "" Then
        MsgBox "请确认机组品牌!"
        comJzPb.SetFocus
        Exit Sub
    End If
    If txtJzxh.Text = "" Then
        MsgBox "请确认机组型号!"
        txtJzxh.SetFocus
        Exit Sub
    End If
    If Val(txtSL.Text) = 0 Then
        MsgBox "请确认数量!"
        txtSL.SetFocus
        Exit Sub
    End If
    If FmxcXJ.txtDrq.Text = "" Then
        FmxcXJ.txtDrq.Text = mod1.DQda
    End If
    If FmxcXJ.txtBrq.Text = "" Then
        FmxcXJ.txtBrq.Text = mod1.DQda
    End If
    
    
    dtgHP.Col = 1: GoodsCode = dtgHP.Text
    dtgHP.Col = 2: MC = Left(dtgHP.Text, 50)
    dtgHP.Col = 3: GG = Left(dtgHP.Text, 50)
    dtgHP.Col = 4: Lb = dtgHP.Text
    dtgHP.Col = 5: Dw = dtgHP.Text
    dtgHP.Col = 6: GoodsId = dtgHP.Text
    If Val(GoodsId) = 0 Then Exit Sub
                                       '新版本速达
        timZm = 2
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "询价单2011"
        mod1.cmd.Parameters("@NBLX") = "速达配件添加"
        mod1.cmd.Parameters("@bh") = ""
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = FmxcXJ.lblBid.ToolTipText
        mod1.cmd.Parameters("@mt2") = FmxcXJ.lblZl.Caption
        mod1.cmd.Parameters("@mt3") = comJzPb.Text  '机组品牌
        mod1.cmd.Parameters("@mt4") = txtJzxh.Text  '机组型号
        mod1.cmd.Parameters("@mt5") = Dw '压缩机型号,单位
        mod1.cmd.Parameters("@mt6") = "" '出厂编号
        mod1.cmd.Parameters("@mt7") = GoodsId '机组序列号
        mod1.cmd.Parameters("@mt8") = MC '零件名称
        mod1.cmd.Parameters("@mt9") = GoodsCode '零件规格号
        mod1.cmd.Parameters("@mt10") = GG '品牌及产地
        mod1.cmd.Parameters("@mt11") = FmxcXJ.txtLx.Text  '业务类型
        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1") = Val(txtSL.Text) '数量
        mod1.cmd.Parameters("@mb1") = 0
        mod1.cmd.Parameters("@md1") = Null
        Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "成功" Then
            MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
            Exit Sub
        Else '提交成功,等待系统中心处理数据
            cmdAdd.Enabled = False
            cmdJG.Enabled = False
            Me.Enabled = False
            frmWaitA.Visible = True
            frmWaitA.Timer2.Enabled = False
    
            frmWaitA.ZOrder 0
            frmWaitA.Timer2.Enabled = True
            timWait.Enabled = True
        End If
    
    
        Set mod1.cmd = Nothing




End If

End Sub

Private Sub cmdGB_Click()
Me.Visible = False
End Sub

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

Public Sub dtgFF()
dtgHP.Clear: dtgN.Clear
dtgHP.Row = 0: dtgHP.Col = 1: dtgHP.Text = "货品编码"
dtgHP.Col = 2: dtgHP.Text = "货品名称"
dtgHP.Col = 3: dtgHP.Text = "规格"
dtgHP.Col = 4: dtgHP.Text = "货品类别"
dtgHP.Col = 5: dtgHP.Text = "单位"
End Sub

Public Sub CX(CT As String)
Dim ii As Integer: Dim oo As Integer: Dim Oi As Integer
Dim Ra: Dim La
Dim Tid As String: Dim Laid As String
On Error Resume Next
CT = CT & "  order by dbo.l_goods.goodsid,dbo.l_goodsunit.unitname desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open CT, mod1.workSD, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
La = UBound(Ra, 2) + 1
frmGxbjSD.dtgHP.Rows = La + 30
frmGxbjSD.dtgN.Rows = La + 30
frmGxbjSD.dtgN.Cols = frmGxbjSD.dtgHP.Cols
mod1.HTP.Close
Set mod1.HTP = Nothing
Call dtgFF
dtgHP.Visible = False
If La = 0 Then
    dtgHP.Visible = True
    Exit Sub
End If
On Error GoTo GXBJSD
'先复制进内表
For oo = 1 To La
    frmGxbjSD.dtgN.Row = oo
    For ii = 1 To 6
        frmGxbjSD.dtgN.Col = ii
        If IsNull(Ra(ii - 1, oo - 1)) = False Then
            frmGxbjSD.dtgN.Text = Ra(ii - 1, oo - 1)
        Else
            frmGxbjSD.dtgN.Text = ""
        End If
    Next
Next
Tid = "": Laid = ""
'再将筛选重复的,有单位的记录
dtgHP.Row = 1
For oo = 1 To La
    dtgN.Row = oo
    dtgN.Col = 1
    If dtgN.Text = Laid Then
        oo = oo + 1
    Else
        Laid = dtgN.Text
        For ii = 1 To 6
            dtgN.Col = ii: dtgHP.Col = ii
            dtgHP.Text = dtgN.Text
        Next
        dtgHP.Row = dtgHP.Row + 1
    End If
Next
dtgHP.Visible = True
frmGxbjSD.Show
frmGxbjSD.ZOrder 0

Exit Sub
GXBJSD:
MsgBox "ok" & oo
End Sub

Private Sub timQuit_Timer()
Dim tt As String
Dim Rb
Dim Lb As Integer
On Error Resume Next
Dim oo As Integer
Dim jj As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then       '添加配件
    frmGXBj.dtgMa.Visible = False
    frmGXBj.adoGx.Requery
                frmGXBj.dtgMa.FixedCols = 1
    Set frmGXBj.dtgMa.DataSource = frmGXBj.adoGx
    '显示商务支持添加的产品（变色）
    For oo = 1 To frmGXBj.dtgMa.Rows
        frmGXBj.dtgMa.Col = 28
        frmGXBj.dtgMa.Row = oo
        If frmGXBj.dtgMa.Text = "True" Then
            For jj = 1 To 25
                frmGXBj.dtgMa.Col = jj
                frmGXBj.dtgMa.CellForeColor = &H8000000D
            Next
        End If
    Next

    If mod1.Bm = "配送中心" And timZm = 1 Then '让配送中心人可以签字
''''''        frmGXBj.lblQM(0).Caption = ""
''''''        frmGXBj.lblQM(1).Caption = ""
''''''        frmGXBj.cmdQm(0).Caption = ""
''''''        frmGXBj.cmdQm(1).Caption = ""
''''''        frmGXBj.lblTm(0).Caption = ""
''''''        frmGXBj.lblTm(1).Caption = ""
        frmGXBj.lblLc.Caption = 1
        frmGXBj.lblLcRen.Caption = mod1.DName
        frmGXBj.lblLcUid.Caption = mod1.DHid
    End If
    

    comJzPb.Text = ""
    comJzXh.Text = ""
    txtYxh.Text = ""
    txtCbh.Text = ""
    txtXlh.Text = ""
    txtLjbh.Text = ""
    txtLjmc.Text = ""
    txtCd.Text = ""
    txtDrq.Text = ""
    txtSL.Text = ""
    txtMj.Text = ""
    txtDj.Text = ""
    txtBrq.Text = ""
    cmdAdd.Enabled = True
    cmdDel.Enabled = True
    Call frmGXBj.dtgMaFF
    frmGXBj.dtgMa.Visible = True
ElseIf timZm = 2 Then
    tt = "select  ljbh,ljmc+'('+jzpb+' '+pbcd+')',mj,dj,jdj,sl,jhg,drq,zbq,delf,lid,ljmc  from xunjiamx where bid=" & Val(FmxcXJ.lblBid.ToolTipText) & " order by delf desc,lid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdTex
    Rb = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lb = UBound(Rb, 2)
    Call FmxcXJ.dtgBrBound(Rb, Lb)
    
End If
timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mm2,mt1,mt2 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0

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
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdJG.Enabled = False
    End If
    Exit Sub

End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


