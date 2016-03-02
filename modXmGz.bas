Attribute VB_Name = "modXmGz"
Option Explicit
Public Gid As Integer
Public FR As Date '一周的第一天
Public LR As Date '一周的最后一天
Public Ti As Boolean '是否为新添加的记录
Dim adoPwf As Object
Public Sub BGLcBut(Nlb As Integer)
Dim tt As String
Dim oo As Integer
On Error Resume Next
For oo = 10 To 1 Step -1
    Unload frmGzNr.lblTm(oo)
    Unload frmGzNr.cmdQm(oo)
    Unload frmGzNr.lblQM(oo)
Next
    frmGzNr.cmdQm(0).Caption = ""
    frmGzNr.lblTm(0).Caption = ""
    tt = "lcBut(" & Nlb & ")"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    mod1.HTP.MoveFirst
    frmGzNr.cmdQm(0).Caption = ""
    frmGzNr.lblQM(0).Caption = mod1.HTP.Fields("LNR").Value
    frmGzNr.lblTm(0).Caption = ""
    mod1.HTP.MoveNext '第一个数组按钮不用添加,所以,跳到下一记录
    For oo = 1 To mod1.HTP.RecordCount - 1
        Load frmGzNr.lblQM(oo)
        Load frmGzNr.cmdQm(oo)
        Load frmGzNr.lblTm(oo)
        frmGzNr.lblQM(oo).Caption = mod1.HTP.Fields("LNR").Value
        frmGzNr.lblQM(oo).Visible = True
        frmGzNr.lblQM(oo).Left = frmGzNr.lblQM(oo - 1).Left + 1100
        frmGzNr.cmdQm(oo).Caption = ""
        frmGzNr.cmdQm(oo).Visible = True
        frmGzNr.cmdQm(oo).Left = frmGzNr.cmdQm(oo - 1).Left + 1100
        frmGzNr.lblTm(oo).Caption = ""
        frmGzNr.lblTm(oo).Visible = True
        frmGzNr.lblTm(oo).Left = frmGzNr.lblTm(oo - 1).Left + 1100
        mod1.HTP.MoveNext
    Next


        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "QMRZAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@NLb") = Nlb
        mod1.cmd.Parameters("@btz") = mod1.BTZ
        mod1.cmd.Parameters("@QDBH") = frmGzNr.lblGid.Caption
        mod1.cmd.Execute
        Set cmd = Nothing
        

End Sub
Public Function dayWeek(ii As Integer) As String
Select Case ii
Case 1
dayWeek = "日"
Case 2
dayWeek = "一"
Case 3
dayWeek = "二"
Case 4
dayWeek = "三"
Case 5
dayWeek = "四"
Case 6
dayWeek = "五"
Case 7
dayWeek = "六"
End Select
End Function


Public Sub xmAdd() '项目跟踪添加
Dim Pd As Integer
Dim tt As String
On Error Resume Next
tt = "Select * from xmGz where gid=" & frmGzNr.lblGid.Caption
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
If mod1.HTP.RecordCount < 1 Then
    MsgBox "出错,请不要关闭程序,与马晓聪联系。"
    Exit Sub
End If
mod1.HTP.Update "xid", Val(frmGzNr.lblXid.Caption)
mod1.HTP.Update "khQc", frmGzNr.lblXmmc.Caption  '客户全称


mod1.HTP.Update "xmAdr", frmGzNr.lblAdr.Caption '项目地址
mod1.HTP.Update "BfMd", Left(frmGzNr.txtBfMd.Text, 500) '拜访目的
mod1.HTP.Update "BfJg", Left(frmGzNr.txtBfJg.Text, 500) '拜访结果
mod1.HTP.Update "xbCC", Left(frmGzNr.txtXBCC.Text, 500) '下步计划
'mod1.htp.Update "XDBZ", Left(frmGzNr.txtXDBZ.Text, 50) '行动步骤
mod1.HTP.Update "aTime", frmGzNr.lblRq.Caption '时间安排
mod1.HTP.Update "XmFy", Val(frmGzNr.txtXmFy.Text)  '项目费用
mod1.HTP.Update "xM", Left(frmGzNr.txtXm.Text, 500) '项目描述
mod1.HTP.Update "jzDC", Left(frmGzNr.txtjzDC.Text, 500) '竞争对手
mod1.HTP.Update "zgPd", Left(frmGzNr.txtzgPd.Text, 500) '主管评定
mod1.HTP.Update "zgQz", frmGzNr.lblZGQZ.Caption '主管签字
mod1.HTP.Update "Lb", 1 '类别

mod1.HTP.UpdateBatch

'保存费用表
frmGzNr.adoFy.Recordset.UpdateBatch

'客户平台
If frmGzNr.optA.Value = True Then
    Pd = 0
ElseIf frmGzNr.optB.Value = True Then
    Pd = 30
ElseIf frmGzNr.optC.Value = True Then
    Pd = 60
ElseIf frmGzNr.optD.Value = True Then
    Pd = 90
End If

'更新项目资料表中的项目平台
tt = "update xmzl set khJb=" & Pd & " where xid=" & Val(frmGzNr.lblXid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
End Sub

Public Sub jiAdd() '计划添加
Dim tt As String
On Error Resume Next
tt = "Select * from xmGz where gid=" & modXmGz.Gid
frmGzJ.adoXmgz.Recordset.Close
frmGzJ.adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
frmGzJ.adoXmgz.Recordset.Update "khDh", frmGzJ.lblDm.Caption '客户代号
frmGzJ.adoXmgz.Recordset.Update "khQc", frmGzJ.lblKhmc.Caption   '客户全称
frmGzJ.adoXmgz.Recordset.Update "xmmc", frmGzJ.lblKhmc.Caption
'客户平台
If frmGzJ.optA.Value = True Then
frmGzJ.adoXmgz.Recordset.Update "khJb", 0
ElseIf frmGzJ.optB.Value = True Then
frmGzJ.adoXmgz.Recordset.Update "khJb", 30
ElseIf frmGzJ.optC.Value = True Then
frmGzJ.adoXmgz.Recordset.Update "khJb", 60
ElseIf frmGzJ.optD.Value = True Then
frmGzJ.adoXmgz.Recordset.Update "khJb", 90
End If

frmGzJ.adoXmgz.Recordset.Update "xmAdr", frmGzJ.lblAdr.Caption '项目地址
frmGzJ.adoXmgz.Recordset.Update "BfMd", Left(frmGzJ.txtBfMd.Text, 500) '拜访目的
'frmGzNr.adoXmgz.Recordset.Update "BfJg", Left(frmGzNr.txtBfJg.Text, 50) '拜访结果
frmGzJ.adoXmgz.Recordset.Update "XDBZ", Left(frmGzJ.txtXDBZ.Text, 500) '行动步骤
frmGzJ.adoXmgz.Recordset.Update "aTime", frmGzJ.lblRq.Caption '时间安排
'frmgzJ.adoXmgz.Recordset.Update "XmFy", Val(frmgzJ.txtXmFy.Text)  '项目费用
frmGzJ.adoXmgz.Recordset.Update "xM", Left(frmGzJ.txtXm.Text, 500) '项目描述
frmGzJ.adoXmgz.Recordset.Update "jzDC", Left(frmGzJ.txtjzDC.Text, 500) '竞争对手
frmGzJ.adoXmgz.Recordset.Update "zgPd", Left(frmGzJ.txtzgPd.Text, 500) '主管评定
frmGzJ.adoXmgz.Recordset.Update "zgQz", frmGzJ.lblZGQZ.Caption '主管签字
frmGzJ.adoXmgz.Recordset.Update "Lb", 0 '类别
frmGzJ.adoXmgz.Recordset.UpdateBatch

''保存费用表
'frmGzNr.adoFy.Recordset.UpdateBatch
End Sub



Public Sub xmBound(Gid As Long)  '项目跟踪绑定
Dim tt As String
On Error Resume Next
tt = "Select * from xmGz where gid=" & Gid
frmGzNr.adoXmgz.Recordset.Close
frmGzNr.adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'frmGzNr.lblDm.Caption = frmGzNr.adoXmgz.Recordset.Fields("khDh").Value '客户代号
frmGzNr.lblXid.Caption = frmGzNr.adoXmgz.Recordset.Fields("xid").Value
frmGzNr.lblXmmc.Caption = frmGzNr.adoXmgz.Recordset.Fields("xmmc").Value '项目全称
frmGzNr.comKhmc.Text = frmGzNr.adoXmgz.Recordset.Fields("khqc").Value '客户名称
frmGzNr.lblYwy.Caption = frmGzNr.adoXmgz.Recordset.Fields("ywy").Value '业务员
frmGzNr.lblGid.Caption = frmGzNr.adoXmgz.Recordset.Fields("gid").Value




frmGzNr.lblAdr.Caption = frmGzNr.adoXmgz.Recordset.Fields("xmAdr").Value '项目地址
frmGzNr.txtBfMd.Text = frmGzNr.adoXmgz.Recordset.Fields("BfMd").Value '拜访目的
frmGzNr.txtBfJg.Text = frmGzNr.adoXmgz.Recordset.Fields("BfJg").Value '拜访结果
frmGzNr.txtXBCC.Text = frmGzNr.adoXmgz.Recordset.Fields("xbCC").Value '下步计划
'frmGzNr.txtXDBZ.Text = frmGzNr.adoXmgz.Recordset.Fields("XDBZ").Value '行动步骤
frmGzNr.lblRq.Caption = frmGzNr.adoXmgz.Recordset.Fields("aTime").Value '时间安排
frmGzNr.txtXmFy.Text = frmGzNr.adoXmgz.Recordset.Fields("XmFy").Value '项目费用
frmGzNr.txtXm.Text = frmGzNr.adoXmgz.Recordset.Fields("xM").Value '项目描述
frmGzNr.txtjzDC.Text = frmGzNr.adoXmgz.Recordset.Fields("jzDC").Value '竞争对手
frmGzNr.txtzgPd.Text = frmGzNr.adoXmgz.Recordset.Fields("zgPd").Value '主管评定
frmGzNr.lblZGQZ.Caption = frmGzNr.adoXmgz.Recordset.Fields("zgQz").Value '主管签字

frmGzNr.lblLc.Caption = frmGzNr.adoXmgz.Recordset.Fields("Lc").Value
frmGzNr.lblLcRen.Caption = Trim(frmGzNr.adoXmgz.Recordset.Fields("LcRen").Value)
frmGzNr.lblLcUid.Caption = Trim(frmGzNr.adoXmgz.Recordset.Fields("LcUid").Value)
frmGzNr.lblFwid.Caption = frmGzNr.adoXmgz.Recordset.Fields("Fwid").Value
frmGzNr.lblNlb.Caption = frmGzNr.adoXmgz.Recordset.Fields("Nlb").Value
frmGzNr.lblLcou.Caption = frmGzNr.adoXmgz.Recordset.Fields("Lcou").Value
frmGzNr.lblHtbh.Caption = frmGzNr.adoXmgz.Recordset.Fields("htbh").Value

'绑定费用表
tt = "Select * from fyTg where gid=" & Val(frmGzNr.lblGid.Caption)
frmGzNr.adoFy.Recordset.Close
frmGzNr.adoFy.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Set frmGzNr.dtgFy.DataSource = frmGzNr.adoFy

'更新客户交往表和内容
    tt = "select ren,llid from xmren where gid=" & Val(frmGzNr.lblGid.Caption) & " order by llid desc"
    frmGzNr.adoBlx.Close
    frmGzNr.adoBlx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGzNr.dtgRen.DataSource = frmGzNr.adoBlx
'    tt = "select ren,llid from xmren where gid=" & Val(lblGid.Caption) & " order by llid desc"
'    adoBlx.Close
'    adoBlx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    frmGzNr.adoBlx.MoveFirst
    tt = "select Tnr,ren from xmRen where llid=" & frmGzNr.adoBlx.Fields("llid").Value
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    frmGzNr.txtJw.Text = mod1.HTP.Fields("tnr").Value
    frmGzNr.lblRen.Caption = mod1.HTP.Fields("ren").Value
    frmGzNr.dtgRen.ColWidth(2) = 0
    frmGzNr.dtgRen.ColWidth(3) = 0
    frmGzNr.dtgRen.ColWidth(4) = 0
    
'取得项目总费用
'计算项目总费用
tt = "select sum(xg) from fybx where khmc='" & frmGzNr.lblXmmc.Caption & "' and not(qrq is null)"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
frmGzNr.lblCxmFy.Caption = mod1.HTP.Fields(0).Value

        frmGzNr.optC.Enabled = False
        frmGzNr.optD.Enabled = False
'客户平台
If mod1.HTP.Fields("khJb").Value = 0 Then
    frmGzNr.optA.Value = True
    Call modXmGz.XMPwf
ElseIf mod1.HTP.Fields("khJb").Value = 30 Then
    frmGzNr.optB.Value = True
    Call modXmGz.XMPwf
ElseIf mod1.HTP.Fields("khJb").Value = 60 Then
    frmGzNr.cmdBJ.Visible = True
    frmGzNr.optC.Value = True
ElseIf mod1.HTP.Fields("khJb").Value = 90 Then
    frmGzNr.cmdBJ.Visible = True
    frmGzNr.optD.Value = True
End If
    
    If frmGzNr.lblLcRen.Caption = mod1.DName And frmGzNr.lblLcUid.Caption = mod1.DHid And Val(frmGzNr.lblLc.Caption) <= Val(frmGzNr.lblLcou.Caption) Then
        frmGzNr.cmdMod.Enabled = True
    Else
        frmGzNr.cmdMod.Enabled = False
    End If
    
    Call modXmGz.OpenXMGZAN(True) '打开按钮
    frmGzNr.cmdSave.Enabled = False
End Sub
Public Sub OpenXMGZAN(LX As Boolean)
Dim tt As String
Dim oo As Integer
On Error Resume Next
If LX = True Then   '销售日记
    For oo = 10 To 1 Step -1
        Unload frmGzNr.cmdQm(oo)
        Unload frmGzNr.lblQM(oo)
        Unload frmGzNr.lblTm(oo)
    Next
    frmGzNr.cmdQm(0).Caption = ""
    frmGzNr.lblTm(0).Caption = ""
      tt = "qmrzOpen(" & mod1.BTZ & ",'" & frmGzNr.lblGid.Caption & "')"
      Set mod1.HTP = CreateObject("adodb.recordset")
      mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
      If mod1.HTP.RecordCount > 0 Then
         mod1.HTP.MoveFirst
         frmGzNr.cmdQm(0).Visible = True
         frmGzNr.lblQM(0).Visible = True
         frmGzNr.lblTm(0).Visible = True
         frmGzNr.cmdQm(0).Caption = mod1.HTP.Fields("Qren").Value
         frmGzNr.lblQM(0).Caption = mod1.HTP.Fields("QLabel").Value
         frmGzNr.lblTm(0).Caption = mod1.HTP.Fields("QRQ").Value
         frmGzNr.cmdQm(0).Tag = mod1.HTP.Fields("zid").Value
         mod1.HTP.MoveNext
         For oo = 1 To mod1.HTP.RecordCount - 1
           Load frmGzNr.lblQM(oo)
           frmGzNr.lblQM(oo).Caption = ""
           Load frmGzNr.cmdQm(oo)
           frmGzNr.cmdQm(oo).Caption = ""
           Load frmGzNr.lblTm(oo)
           frmGzNr.lblTm(oo).Caption = ""
           frmGzNr.lblQM(oo).Caption = mod1.HTP.Fields("QLabel").Value
           frmGzNr.cmdQm(oo).Caption = mod1.HTP.Fields("Qren").Value
           frmGzNr.lblTm(oo).Caption = mod1.HTP.Fields("QRQ").Value
           frmGzNr.cmdQm(oo).Tag = mod1.HTP.Fields("zid").Value
           frmGzNr.lblQM(oo).Visible = True
           frmGzNr.cmdQm(oo).Visible = True
           frmGzNr.lblTm(oo).Visible = True
           frmGzNr.lblQM(oo).Left = frmGzNr.lblQM(oo - 1).Left + 1100
           frmGzNr.cmdQm(oo).Left = frmGzNr.cmdQm(oo - 1).Left + 1100
           frmGzNr.lblTm(oo).Left = frmGzNr.lblTm(oo - 1).Left + 1100
           mod1.HTP.MoveNext
        Next
     Else
        frmGzNr.cmdQm(0).Visible = False
        frmGzNr.lblQM(0).Visible = False
        frmGzNr.lblTm(0).Visible = False
     End If
End If
End Sub
Public Sub jiBound() '计划绑定
Dim tt As String
On Error Resume Next
tt = "Select * from xmGz where gid=" & modXmGz.Gid
frmGzJ.adoXmgz.Recordset.Close
frmGzJ.adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
frmGzJ.lblDm.Caption = frmGzJ.adoXmgz.Recordset.Fields("khDh").Value '客户代号
frmGzJ.lblKhmc.Caption = frmGzJ.adoXmgz.Recordset.Fields("xmmc").Value '客户全称
frmGzJ.lblYwy.Caption = frmGzJ.adoXmgz.Recordset.Fields("ywy").Value '业务员

'客户平台
If frmGzJ.adoXmgz.Recordset.Fields("khJb").Value = 0 Then
    frmGzJ.optA.Value = True
ElseIf frmGzJ.adoXmgz.Recordset.Fields("khJb").Value = 30 Then
    frmGzJ.optB.Value = True
ElseIf frmGzJ.adoXmgz.Recordset.Fields("khJb").Value = 60 Then
    frmGzJ.optC.Value = True
ElseIf frmGzJ.adoXmgz.Recordset.Fields("khJb").Value = 90 Then
    frmGzJ.optD.Value = True
End If

frmGzJ.lblAdr.Caption = frmGzJ.adoXmgz.Recordset.Fields("xmAdr").Value '项目地址
frmGzJ.txtBfMd.Text = frmGzJ.adoXmgz.Recordset.Fields("BfMd").Value '拜访目的
'frmgzJ.txtBfJg.Text = frmgzJ.adoXmgz.Recordset.Fields("BfJg").Value '拜访结果
'frmgzJ.txtXBCC.Text = frmgzJ.adoXmgz.Recordset.Fields("xbCC").Value '下步计划
frmGzJ.txtXDBZ.Text = frmGzJ.adoXmgz.Recordset.Fields("XDBZ").Value '行动步骤
frmGzJ.lblRq.Caption = frmGzJ.adoXmgz.Recordset.Fields("aTime").Value '时间安排
'frmGzJ.txtXmFy.Text = frmGzJ.adoXmgz.Recordset.Fields("XmFy").Value '项目费用
frmGzJ.txtXm.Text = frmGzJ.adoXmgz.Recordset.Fields("xM").Value '项目描述
frmGzJ.txtjzDC.Text = frmGzJ.adoXmgz.Recordset.Fields("jzDC").Value '竞争对手
frmGzJ.txtzgPd.Text = frmGzJ.adoXmgz.Recordset.Fields("zgPd").Value '主管评定
frmGzJ.lblZGQZ.Caption = frmGzJ.adoXmgz.Recordset.Fields("zgQz").Value '主管签字

''绑定费用表
'tt = "Select * from fyTg where gid=" & modXmGz.Gid
'frmGzJ.adoFy.Recordset.Close
'frmGzJ.adoFy.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'Set frmGzJ.dtgFy.DataSource = frmGzJ.adoFy

End Sub

Public Sub xmQing() '项目跟踪字段清空
frmGzNr.lblXid.Caption = "" '客户代号
frmGzNr.lblXmmc.Caption = "" '客户全称
frmGzNr.comKhmc.Text = ""
frmGzNr.lblAdr.Caption = ""
frmGzNr.lblKid.Caption = ""
frmGzNr.lblGid.Caption = ""
frmGzNr.lblYwy.Caption = ""

'客户平台
    frmGzNr.optA.Value = False
    frmGzNr.optB.Value = False
    frmGzNr.optC.Value = False
    frmGzNr.optD.Value = False
frmGzNr.txtJw.Text = "" '客户交往情况
frmGzNr.lblAdr.Caption = "" '项目地址
frmGzNr.txtBfMd.Text = ""  '拜访目的
frmGzNr.txtBfJg.Text = ""  '拜访结果
'frmGzNr.txtXDBZ.Text = "" '行动步骤
frmGzNr.txtXBCC.Text = "" '下步措施
frmGzNr.lblRq.Caption = ""  '时间安排
frmGzNr.txtXmFy.Text = "" '项目费用
frmGzNr.lblCxmFy.Caption = "" '项目总费用
frmGzNr.txtXm.Text = ""  '项目描述
frmGzNr.txtjzDC.Text = ""  '竞争对手
frmGzNr.txtzgPd.Text = ""  '主管评定
frmGzNr.lblZGQZ.Caption = ""  '主管签字

Set frmGzNr.dtgRen.DataSource = Nothing
frmGzNr.comRen.Text = ""
frmGzNr.lblRen.Caption = ""
frmGzNr.cmdRenAdd.Visible = False
frmGzNr.cmdRenDel.Visible = False
frmGzNr.cmdFadd.Visible = False
frmGzNr.cmdFdel.Visible = False
frmGzNr.cmdTg.Visible = False
frmGzNr.cmdBJ.Visible = False
frmGzNr.cmdRenAdd.Visible = False

frmGzNr.lblLc.Caption = ""
frmGzNr.lblLcRen.Caption = ""
frmGzNr.lblLcUid.Caption = ""
frmGzNr.lblFwid.Caption = ""
frmGzNr.lblNlb.Caption = ""
frmGzNr.lblHtbh.Caption = "" '本次报告费用所归的合同编号
Set frmGzNr.dtgRen.DataSource = Nothing

End Sub

Public Sub jhQing() '计划字段清空
frmGzJ.lblDm.Caption = "" '客户代号
'frmGzJ.lblXmmc.Caption = "" '客户全称

'客户平台
    frmGzJ.optA.Value = False
    frmGzJ.optB.Value = False
    frmGzJ.optC.Value = False
    frmGzJ.optD.Value = False

frmGzJ.lblAdr.Caption = "" '项目地址
frmGzJ.txtBfMd.Text = ""  '拜访目的
'frmGzJ.txtBfJg.Text = ""  '拜访结果
frmGzJ.txtXDBZ.Text = "" '行动步骤
frmGzJ.lblRq.Caption = ""  '时间安排
'frmGzJ.txtXmFy.Text = "" '项目费用
frmGzJ.txtXm.Text = ""  '项目描述
frmGzJ.txtjzDC.Text = ""  '竞争对手
frmGzJ.txtzgPd.Text = ""  '主管评定
frmGzJ.lblZGQZ.Caption = ""  '主管签字

End Sub



Public Sub FyQing() '营销部报销单清空

    frmFYBX.lblBh.Caption = ""
    frmFYBX.comQy.Caption = "上海"
    frmFYBX.txtHg.Text = ""
    frmFYBX.lblDx.Caption = ""
    frmFYBX.lblFR.Caption = ""
    frmFYBX.lblLR.Caption = ""
    frmFYBX.lblRq.Caption = ""
    frmFYBX.cmdBxr.Caption = ""
    frmFYBX.cmdJc.Caption = ""
    frmFYBX.cmdJl.Caption = ""
    frmFYBX.cmdZj.Caption = ""
    frmFYBX.txtQc.Text = ""
    frmFYBX.txtCwBZ.Text = ""
    frmFYBX.txtBz.Text = ""
    frmFYBX.lblTa.Caption = ""
    frmFYBX.lblTb.Caption = ""
    frmFYBX.lblTC.Caption = ""
    frmFYBX.lblTd.Caption = ""
    frmFYBX.lblNlb.Caption = ""
    frmFYBX.cmdJc.Tag = ""
    frmFYBX.cmdJl.Tag = ""
    frmFYBX.cmdZj.Tag = ""
    frmFYBX.txtQc.Tag = ""
    
End Sub








Public Sub XMPwf() '项目评审合格否
Dim tt As String
Dim Pwf As Boolean
Pwf = True
On Error Resume Next
Set adoPwf = CreateObject("adodb.recordset")
tt = "Select * from XmPwf where xid=" & Val(frmGzNr.lblXid.Caption)
adoPwf.Close
adoPwf.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
adoPwf.MoveFirst
Do While Not adoPwf.EOF
    If adoPwf.Fields("pwf").Value = False Then
        Pwf = False
        Exit Do
    End If
    adoPwf.MoveNext
Loop

If Pwf = True Then
    frmGzNr.optC.Enabled = True
    frmGzNr.optD.Enabled = True
Else
    frmGzNr.optC.Enabled = False
    frmGzNr.optD.Enabled = False
End If
End Sub
