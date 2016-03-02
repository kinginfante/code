VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form FmxcLxNew 
   BackColor       =   &H00FFFFC0&
   Caption         =   "业务类型选择"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   10170
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   5040
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2820
      Top             =   6090
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgNewLx 
      Height          =   5925
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   10451
      _Version        =   393216
      BackColor       =   16777152
      Rows            =   14
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin NiceFormControl.NiceButton cmdNew 
      Height          =   345
      Left            =   4170
      TabIndex        =   1
      Top             =   6120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   609
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "fmxcLxNew.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Caption         =   "生成询价单"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请用鼠标单击相应的栏目"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   6210
      Width           =   2535
   End
End
Attribute VB_Name = "FmxcLxNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LX As String
Public Lc As Single  '选择的业务序号
Public Hid As Integer
Dim Bid As Long
Dim timZm As Integer

Private Sub cmdNew_Click()
If Lc = 0 Then
    MsgBox "请选择相应的业务类型!(在列表中双击)"
    Exit Sub
End If
If Hid = 0 And (Lc = 20 Or Lc = 21 Or Lc = 22 Or Lc = 23) Then
    Exit Sub
End If
If Me.Hid = 0 Then
    Call FMXCXmmc.Qing
    FMXCXmmc.Show
    FMXCXmmc.ZOrder 0
    FMXCXmmc.Lb = "询价单"
    FMXCXmmc.NiceButton1.Caption = "生 成 单 据 (询价单)"
Else
'''    FMXCXmmc.Show
'''    'Call FMXCXmmc.NiceButton1_Click
    If Right(cmdNew.Caption, 5) <> "成本变更单" Then
        timZm = 2
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.workKK
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "新合同2013"
        mod1.cmd.Parameters("@NBLX") = "添加询价单"
        mod1.cmd.Parameters("@bh") = ""
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = FmxcNew.txtXmmc.Text
        mod1.cmd.Parameters("@mt2") = FmxcLxNew.LX 'ZL
        mod1.cmd.Parameters("@mt5") = FmxcNew.txtKhmc.Text
        mod1.cmd.Parameters("@mt25") = FmxcLxNew.Hid
        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1") = Val(FmxcNew.txtXmmc.ToolTipText)
        mod1.cmd.Parameters("@mm2") = Val(FmxcLxNew.cmdNew.ToolTipText)
        dtgNewLx.Col = 7
        mod1.cmd.Parameters("@mb1") = dtgNewLx
        'Exit Sub
        mod1.cmd.Parameters("@md1") = Null
        Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute
       ' MsgBox "b"
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "成功" Then
            MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
            If timZm = 1 Then
                cmdNew.Enabled = False
            End If
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
    Else
        Dim YZE As Single
        Dim XZE As Single
        '''''If FmxcNew.Visible = True Then '检验是否超出预估成本总额
        '''''    YZE = Val(FmxcNew.txtQb.Text)
        '''''    dtgN.Col = 2
        '''''    dtgN
        '''''End If
        Call fmxcZJ.Qing

        '引用合同数据
            fmxcZJ.lblKhmc.Caption = FmxcNew.txtKhmc.Text
            fmxcZJ.lblGLBH.Caption = FmxcNew.txtHtbh.Text
            fmxcZJ.lblGLBH.ToolTipText = FmxcNew.lblHid.Caption
            fmxcZJ.lblZbh.Caption = FmxcNew.txtZbh.Text
            fmxcZJ.lblXz.Caption = FmxcNew.txtHtxz.Text
            fmxcZJ.lblZe.Caption = FmxcNew.txtHtze.Text
            fmxcZJ.lblFBF.Caption = LX
            fmxcZJ.htRow = Lc
        fmxcZJ.cmdSave.Enabled = True
        
        fmxcZJ.lblYwy.Caption = mod1.DName
        Call fmxcZJ.dtgFF1
        Call fmxcZJ.dtgPFF
        fmxcZJ.Show
        fmxcZJ.ZOrder 0
        Me.Visible = False
        fmxcZJ.frmGui.Visible = True
        fmxcZJ.frmGui.Enabled = True
        fmxcZJ.comGui.Enabled = True
        fmxcZJ.optF.Enabled = True
        fmxcZJ.optQ.Enabled = True
    
    End If
End If

End Sub

Private Sub dtgNewLx_Click()
Dim L1 As String
Dim L2 As String
Dim L3 As String
Dim L0 As String
'MsgBox dtgNewLx.Row & " " & dtgNewLx.Col
Dim Lrow As Integer
Dim oo As Integer
Lrow = dtgNewLx.Row
On Error Resume Next


dtgNewLx.Visible = False
'变颜色
For oo = 1 To 50
    dtgNewLx.Row = oo
    dtgNewLx.Col = 0: dtgNewLx.CellForeColor = &H0&
    dtgNewLx.Col = 1: dtgNewLx.CellForeColor = &H0&
    dtgNewLx.Col = 2: dtgNewLx.CellForeColor = &H0&
    dtgNewLx.Col = 3: dtgNewLx.CellForeColor = &H0&
    dtgNewLx.Col = 4: dtgNewLx.CellForeColor = &H0&
Next
'dtgNewLx.ForeColor = &H0&

dtgNewLx.Row = Lrow
    dtgNewLx.Col = 3: L3 = dtgNewLx.Text
    dtgNewLx.Col = 2: L2 = dtgNewLx.Text
    dtgNewLx.Col = 1: L1 = dtgNewLx.Text
    If L3 <> L2 Then
        dtgNewLx.Col = 3: dtgNewLx.CellForeColor = &HFF&
    ElseIf L3 = L2 And L2 <> L1 Then
        dtgNewLx.Col = 2: dtgNewLx.CellForeColor = &HFF&
    ElseIf L3 = L2 And L2 = L1 And Trim(L1) <> "" Then
        dtgNewLx.Col = 1: dtgNewLx.CellForeColor = &HFF&
    ElseIf L3 = L2 And L2 = L1 And Trim(L1) = "" Then
        dtgNewLx.Col = 0: dtgNewLx.CellForeColor = &HFF&
    End If
    dtgNewLx.Col = 4: dtgNewLx.CellForeColor = &HFF&
dtgNewLx.Visible = True

If dtgNewLx.Row = 0 Then Exit Sub
dtgNewLx.Col = 6
Lc = dtgNewLx.Text
dtgNewLx.Col = 3: LX = dtgNewLx.Text
If InStr(1, LX, "人工") > 0 Then
    dtgNewLx.Col = 1
    LX = dtgNewLx.Text & "->" & LX
End If
If Right(cmdNew.Caption, 3) = "询价单" Then
    cmdNew.Caption = "生成" & LX & "询价单"
    If Trim(LX) = "" Then
        LX = "其他（非材料）"
        cmdNew.Caption = "生成其他(非材料)询价单"
    End If
Else
    cmdNew.Caption = "生成" & LX & "成本变更单"
End If
cmdNew.ToolTipText = Lc
End Sub

Private Sub dtgNewLx_DblClick()
Dim LT As String
dtgNewLx.Col = 8
LT = dtgNewLx.Text
If dtgNewLx.Text = "" Then Exit Sub
FmxcXJ.txtLx.Text = LT

'Me.Visible = False
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Width = 10000
Me.Height = 6990
Call Me.NewLx
End Sub
Public Sub NewLx()
Dim tt As String
Dim Ra
Dim La As Integer
Dim oo As Integer
tt = "select la,lb,lc,ld,le,lf,zid,lx,xc from newLx order by zid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgNewLx.Rows = La + 20
dtgNewLx.Cols = 9
dtgNewLx.Clear
dtgNewLx.Refresh
On Error Resume Next
For oo = 0 To La - 1
    dtgNewLx.Row = oo
    dtgNewLx.Col = 0: dtgNewLx.Text = Ra(0, oo)
    dtgNewLx.Col = 1: dtgNewLx.Text = Ra(1, oo)
    dtgNewLx.Col = 2: dtgNewLx.Text = Ra(2, oo)
    dtgNewLx.Col = 3: dtgNewLx.Text = Ra(3, oo)
    dtgNewLx.Col = 4: dtgNewLx.Text = Ra(4, oo)
    dtgNewLx.Col = 5: dtgNewLx.Text = Ra(5, oo)
    dtgNewLx.Col = 6: dtgNewLx.Text = Ra(6, oo)
    dtgNewLx.Col = 7: dtgNewLx.Text = Ra(7, oo)
    dtgNewLx.Col = 8: dtgNewLx.Text = Ra(8, oo)
Next
'''''For oo = 0 To La - 1
'''''    dtgNewLx.Row = oo
'''''    If IsNull(ra(0, oo)) = True Then ra(0, oo) = " ": dtgNewLx.Col = 0: dtgNewLx.Text = ra(0, oo)
'''''    If IsNull(ra(1, oo)) = True Then ra(1, oo) = " ": dtgNewLx.Col = 1: dtgNewLx.Text = ra(1, oo)
'''''    If IsNull(ra(2, oo)) = True Then ra(2, oo) = " ": dtgNewLx.Col = 2: dtgNewLx.Text = ra(2, oo)
'''''    If IsNull(ra(3, oo)) = True Then ra(3, oo) = " ": dtgNewLx.Col = 3: dtgNewLx.Text = ra(3, oo)
'''''    If IsNull(ra(4, oo)) = True Then ra(4, oo) = " ": dtgNewLx.Col = 4: dtgNewLx.Text = ra(4, oo)
'''''    If IsNull(ra(5, oo)) = True Then ra(5, oo) = " ": dtgNewLx.Col = 5: dtgNewLx.Text = ra(5, oo)
'''''Next
dtgNewLx.ColWidth(0) = 900
dtgNewLx.ColWidth(3) = 1100
dtgNewLx.ColWidth(4) = 5505
dtgNewLx.ColWidth(8) = 0
dtgNewLx.Row = 0
dtgNewLx.MergeCells = flexMergeFree
dtgNewLx.MergeRow(0) = True
dtgNewLx.MergeRow(1) = True
dtgNewLx.MergeRow(2) = True
dtgNewLx.MergeRow(3) = True
dtgNewLx.MergeRow(4) = True
dtgNewLx.MergeRow(5) = True
dtgNewLx.MergeRow(6) = True '
dtgNewLx.MergeRow(7) = True
dtgNewLx.MergeRow(8) = True
dtgNewLx.MergeRow(9) = True
dtgNewLx.MergeRow(10) = True
dtgNewLx.MergeRow(11) = True
dtgNewLx.MergeRow(12) = True
dtgNewLx.MergeRow(13) = True
dtgNewLx.MergeRow(14) = True
dtgNewLx.MergeRow(15) = True
dtgNewLx.MergeRow(16) = True
dtgNewLx.MergeRow(17) = True
dtgNewLx.MergeRow(18) = True
dtgNewLx.MergeRow(19) = True
dtgNewLx.MergeRow(20) = True
dtgNewLx.MergeRow(21) = True
dtgNewLx.MergeRow(22) = True
dtgNewLx.MergeRow(23) = True
dtgNewLx.MergeRow(24) = True
dtgNewLx.MergeRow(25) = True
dtgNewLx.MergeRow(26) = True
dtgNewLx.MergeRow(27) = True
dtgNewLx.MergeRow(28) = True
dtgNewLx.MergeRow(29) = True
dtgNewLx.MergeRow(30) = True
dtgNewLx.MergeRow(31) = True
dtgNewLx.MergeRow(32) = True
dtgNewLx.MergeRow(33) = True
dtgNewLx.MergeRow(34) = True
dtgNewLx.MergeRow(35) = True
dtgNewLx.MergeRow(36) = True
dtgNewLx.MergeRow(37) = True
dtgNewLx.MergeRow(38) = True
dtgNewLx.MergeRow(39) = True
dtgNewLx.MergeRow(40) = True
dtgNewLx.MergeRow(41) = True
dtgNewLx.MergeCol(0) = True
dtgNewLx.MergeCol(1) = True
dtgNewLx.MergeCol(2) = True
dtgNewLx.MergeCol(3) = True
dtgNewLx.MergeCol(4) = True
dtgNewLx.MergeCol(5) = True
dtgNewLx.Row = 0:
dtgNewLx.Col = 0: dtgNewLx.CellFontBold = True
dtgNewLx.Col = 4: dtgNewLx.CellFontBold = True
dtgNewLx.Col = 5: dtgNewLx.CellFontBold = True
dtgNewLx.Refresh
End Sub

Private Sub timQuit_Timer()
Dim htRow As Integer
Dim tt As String
Dim Rf
On Error Resume Next
Dim ii As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '如果为添加合同评审
'''''''    Call FmxcNew.Bound(Hid)
'''''''    FmxcNew.Show
'''''''    FmxcNew.ZOrder 0
'''''''    FmxcNew.txtBz.Visible = False
'''''''    FmxcNew.cmdSave.Enabled = True
'''''''    FmxcNew.optXm.Visible = False
'''''''    FmxcNew.frmFk.Visible = True
'''''''    For ii = 0 To 4
'''''''        FmxcNew.Shape1(ii).Visible = True
'''''''    Next
'''''''    FmxcNew.comFPLX.Visible = True
'''''''    FmxcNew.companyId.Visible = True
'''''''    FmxcNew.dt3.Visible = True
'''''''    FmxcNew.dt4.Visible = True
ElseIf timZm = 2 Then
    Call FmxcXJ.Bound(Bid)
    FmxcXJ.Show
    FmxcXJ.ZOrder 0
    FmxcXJ.cmdSave.Enabled = True
    
    
'旧版本2012
'''''    HtRow = Val(FmxcLx.cmdNew.ToolTipText)\

'新版本2013
htRow = Val(FmxcLxNew.cmdNew.ToolTipText)
    
'新版本2013

    If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 And LX <> "三菱" Or htRow = 8 And LX <> "松下" Or htRow >= 20 Or InStr(1, "维保大修其他人工压缩机维修保养中介业务分包运费吊装费工程人工", LX) > 0 Or _
     (Val(FmxcXJ.lblBid.ToolTipText) >= 20512 And FmxcXJ.lblZl.ToolTipText = True) Then
     
        FmxcXJ.frmWB.Visible = True
    Else
        FmxcXJ.frmSd.Visible = True
    End If
    If FmxcNew.Visible = True Then
        tt = "select zl,jhg,0,'BJD'+cast(bid as nvarchar(20)),lc,0,bid,lcren from xunjiaD where htbh='" & Trim(Str(FmxcNew.lblHid.Caption)) & "' and delf=1 order by bid"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Rf = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        Call FmxcNew.LXBound(Rf, Rg)
    End If
'旧版本2012
'''''''    If HtRow = 1 Or HtRow = 2 Or HtRow = 3 Or HtRow = 4 Or HtRow = 6 Or HtRow = 12 Then
'''''''        FmxcXJ.frmWB.Visible = True
'''''''    Else
'''''''        FmxcXJ.frmSd.Visible = True
'''''''    End If
End If
timQuit.Enabled = False
Hid = 0
Me.Enabled = True
Me.Visible = False
End Sub


Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    mod1.Ti = 5
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    timWait.Enabled = False
    If timZm = 1 Then
        Hid = mod1.WP.Fields("bh").Value
    Else
        Bid = mod1.WP.Fields("bh").Value
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    If timZm = 1 Then
        NiceButton1.Enabled = False
    End If
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        NiceButton1.Enabled = False
    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


