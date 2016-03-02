VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcXjHp 
   BackColor       =   &H00C0FFC0&
   Caption         =   "货品查询"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10200
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   10200
   StartUpPosition =   3  '窗口缺省
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   345
      Left            =   5370
      TabIndex        =   8
      Top             =   5820
      Visible         =   0   'False
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   609
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
   Begin VB.CommandButton cmdTD 
      Caption         =   "查询替代"
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Top             =   5730
      Width           =   1275
   End
   Begin VB.CommandButton cmdGB 
      Caption         =   "关闭"
      Height          =   345
      Left            =   9120
      TabIndex        =   5
      Top             =   5700
      Width           =   1065
   End
   Begin VB.TextBox txtZ 
      Height          =   285
      Left            =   150
      TabIndex        =   4
      Top             =   5730
      Width           =   1905
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "查  询"
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   5730
      Width           =   1275
   End
   Begin VB.TextBox txtSl 
      Height          =   270
      Left            =   6960
      TabIndex        =   2
      Top             =   5730
      Width           =   615
   End
   Begin VB.CommandButton cmdDao 
      BackColor       =   &H00C0C0FF&
      Caption         =   "导入"
      Height          =   345
      Left            =   7950
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5700
      Width           =   1035
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
      FixedCols       =   0
      BackColorFixed  =   12648384
      BackColorBkg    =   12648384
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "数量"
      Height          =   225
      Left            =   6420
      TabIndex        =   6
      Top             =   5760
      Width           =   495
   End
End
Attribute VB_Name = "FmxcXjHp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bh As String
Dim timZm As Integer '1导入
Public Sub dtgHPFF()
dtgHP.Clear
dtgHP.Rows = 100
dtgHP.Cols = 3
dtgHP.Row = 0
dtgHP.Col = 0: dtgHP.Text = "编码": dtgHP.CellFontBold = True
dtgHP.Col = 1: dtgHP.Text = "货品名称": dtgHP.CellFontBold = True
dtgHP.Col = 2: dtgHP.Text = "描述": dtgHP.CellFontBold = True
dtgHP.ColWidth(1) = 2070
dtgHP.ColWidth(2) = 6750

dtgN.Clear
dtgN.Rows = 100
dtgN.Cols = 3

End Sub

Private Sub cmdC_Click()
Dim tt As String
    tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz from nlpmxc where (partname like '%" & _
        txtZ.Text & "%' or oname like '%" & txtZ.Text & "%' or pb like '%" & txtZ.Text & "%' or jz like '%" & txtZ.Text & _
        "%' or bh like '%" & txtZ.Text & "%' or xn like '%" & txtZ.Text & "%') and delf=1 and lc>2 and jyf=1 and npf=1 order by bh desc"
Call Me.Bound(tt)
End Sub

Public Sub Bound(tt As String)
Dim Ra, Rb
Dim La As Long: Dim Lb As Long
Dim oo As Long
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1: Lb = UBound(Rb, 2) + 1
dtgHP.Visible = False
Call dtgHPFF
dtgHP.Rows = La + 50: dtgN.Rows = La + 50
For oo = 1 To La
    dtgHP.Row = oo: dtgHP.RowHeight(oo) = dtgHP.RowHeight(0) * 2
    dtgHP.Col = 0: dtgHP.Text = Ra(0, oo - 1)
    dtgHP.Col = 1: dtgHP.Text = Ra(1, oo - 1)
    dtgHP.Col = 2: dtgHP.Text = Ra(2, oo - 1)
    dtgN.Row = oo: dtgN.RowHeight(oo) = dtgN.RowHeight(0) * 2
    dtgN.Col = 0: dtgN.Text = Ra(0, oo - 1)
    dtgN.Col = 1: dtgN.Text = Ra(1, oo - 1)
    dtgN.Col = 2: dtgN.Text = Ra(2, oo - 1)
Next
For oo = La + 1 To La + Lb
    dtgHP.Row = oo: dtgHP.RowHeight(oo) = dtgHP.RowHeight(0) * 2
    dtgHP.Col = 0: dtgHP.Text = Rb(0, oo - La - 1)
    dtgHP.Col = 1: dtgHP.Text = Rb(1, oo - La - 1)
    dtgHP.Col = 2: dtgHP.Text = Rb(2, oo - La - 1)
    dtgN.Row = oo: dtgN.RowHeight(oo) = dtgN.RowHeight(0) * 2
    dtgN.Col = 0: dtgN.Text = Rb(0, oo - La - 1)
    dtgN.Col = 1: dtgN.Text = Rb(1, oo - La - 1)
    dtgN.Col = 2: dtgN.Text = Rb(2, oo - La - 1)
Next
dtgHP.Visible = True
End Sub

Private Sub Command1_Click()

End Sub


Private Sub cmdDao_Click()
On Error Resume Next

Dim hg As Long
'If comLx.Text = "" Then Exit Sub
    If Val(txtSl.Text) = 0 Then
        MsgBox "请确认数量!"
        txtSl.SetFocus
        Exit Sub
    End If
    If FmxcXJ.txtDRQ.Text = "" Then
        FmxcXJ.txtDRQ.Text = mod1.DQda
    End If
    If FmxcXJ.txtBrq.Text = "" Then
        FmxcXJ.txtBrq.Text = mod1.DQda
    End If


    If Bh = "" Then Exit Sub
    
If FmxcXJ.Visible = True Then
                                       '新版本速达
        timZm = 2
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.CC
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "询价单2011"
        mod1.cmd.Parameters("@NBLX") = "豪曼配件添加"
        mod1.cmd.Parameters("@bh") = ""
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = FmxcXJ.lblBid.ToolTipText
        mod1.cmd.Parameters("@mt2") = FmxcXJ.lblZl.Caption

      mod1.cmd.Parameters("@mt7") = Bh

        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1") = Val(txtSl.Text) '数量
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
ElseIf fmxcZJ.Visible = True Then
                                      
        timZm = 3
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.CC
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "成本追加单"
        mod1.cmd.Parameters("@NBLX") = "新货品添加"
        mod1.cmd.Parameters("@bh") = fmxcZJ.lblZid.ToolTipText
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = ""
        If fmxcZJ.Visible = True And cmdDao.Caption = "分包导入" Then
            mod1.cmd.Parameters("@mt7").Value = "分包"
        Else
            mod1.cmd.Parameters("@mt7") = Bh
        End If
        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1") = Val(txtSl.Text) '数量
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

Private Sub cmdTD_Click()
Dim tt As String
If Val(Bh) = 0 Then Exit Sub
If cmdTD.Caption = "查询替代" Then
    tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz from nlpmxc where bh='" & Bh & "';" & _
    "select bh,partname,detail from nlpmxcTdb where ybh='" & Bh & "' and delf=1 and lc>1 and jyf=1  order by tid desc"
Else
    tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' '+ff+' 适用机组:'+jz from nlpmxc where bh='" & Bh & "';" & _
    "select bh,partname,detail from nlpmxcTd where ybh='" & Bh & "' and delf=1 and lc>1 and jyf=1  order by tid desc"

End If
Call Me.Bound(tt)
End Sub


Private Sub dtgHP_Click()
dtgN.Row = dtgHP.Row
dtgN.Col = 0
Bh = dtgN.Text
If Left(Bh, 1) = "9" Then
    cmdTD.Caption = "查询替代"
Else
    cmdTD.Caption = "查询原厂"
End If
End Sub

Private Sub Form_Load()
timWait.Enabled = False
timQuit.Enabled = False
End Sub

Private Sub timQuit_Timer()
Dim tt As String
Dim Rb, RC, RD, RE
Dim Lb As Integer
On Error Resume Next
Dim oo As Integer
Dim jj As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 2 Then
    tt = "select ljbh,detail,mj,dj,jdj,sl,jhg,drq,zbq,delf,lid,ljmc,gyid1,gyid2,gyid3,gdj1,gdj2,gdj3,mc1,mc2,mc3,gyid  from XJDetail where bid=" & Val(FmxcXJ.lblBid.ToolTipText) & " order by delf desc,lid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rb = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lb = UBound(Rb, 2)
    Call FmxcXJ.dtgBrBound(Rb, Lb)
ElseIf timZm = 3 Then
    tt = "declare @hid int;" & _
        "select @hid=hid from htzui where zid=" & Val(fmxcZJ.lblZid.ToolTipText) & ";" & _
    "select bh,nr,dj,jdj,sl,ze,delf,did,gyid1,gyid2,gyid3,gdj1,gdj2,gdj3,mc1,mc2,mc3,gyid  from zuijiaDetail where zid=" & Val(fmxcZJ.lblZid.ToolTipText) & " order by delf desc,did desc;" & _
            "select sum(ze) from htzuidetail where zid=" & Val(fmxcZJ.lblZid.ToolTipText) & ";" & _
        "select sum(ze) from htzuiZe where hid=@hid"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    RC = mod1.HTP.GetRows
    Set mod1.HTP = mod1.HTP.NextRecordset
    RD = mod1.HTP.GetRows
    Set mod1.HTP = mod1.HTP.NextRecordset
    RE = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Call fmxcZJ.NewMxBound(RC, RD, RE)
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


