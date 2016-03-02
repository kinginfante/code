Attribute VB_Name = "modPld"


Public Sub PLDBound(Pmid As Long)   '配料单绑定
Dim tt As String
On Error Resume Next
frmPld.Visible = True
tt = "PLDBoundA(" & Pmid & ")"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
frmPld.txtXmmc.Text = mod1.HTP.Fields("xmmc").Value
frmPld.txtKhAdr.Text = mod1.HTP.Fields("xmAdr").Value
frmPld.txtHtbh.Text = mod1.HTP.Fields("htbh").Value
frmPld.txtHtze.Text = mod1.HTP.Fields("Htze").Value
frmPld.lblLc.Caption = mod1.HTP.Fields("LC").Value '流程
frmPld.txtZfyy.Text = mod1.HTP.Fields("zfyy").Value '作废原因
frmPld.txtCb.Text = mod1.HTP.Fields("Tze").Value '成本总额
frmPld.txtGdyy.Text = mod1.HTP.Fields("GdYY").Value '改单原因
frmPld.lblXz.Caption = mod1.HTP.Fields("XZ").Value '合同性质
frmPld.lblKDRQ.Caption = mod1.HTP.Fields("KRQ").Value '开单日期

frmPld.cmdQMA.Caption = mod1.HTP.Fields("QMA").Value
frmPld.cmdQMB.Caption = mod1.HTP.Fields("QMB").Value
frmPld.cmdQMC.Caption = mod1.HTP.Fields("QMC").Value
frmPld.cmdQMD.Caption = mod1.HTP.Fields("QMD").Value
frmPld.cmdQME.Caption = mod1.HTP.Fields("QME").Value
frmPld.cmdYw.Caption = mod1.HTP.Fields("QMyw").Value
If frmPld.cmdYw.Caption = "" And frmPld.cmdQMA.Caption <> "" Then
    frmPld.lblQmA.Caption = "技术支持"
Else
    frmPld.lblQmA.Caption = "销售经理"
End If
frmPld.lblYwT.Caption = mod1.HTP.Fields("QMyt").Value
frmPld.lblTa.Caption = mod1.HTP.Fields("QMAT").Value
frmPld.lblTb.Caption = mod1.HTP.Fields("QMBT").Value
frmPld.lblTC.Caption = mod1.HTP.Fields("QMCT").Value
frmPld.lblTd.Caption = mod1.HTP.Fields("QMDT").Value
frmPld.lblTe.Caption = mod1.HTP.Fields("QMET").Value
frmPld.lblBe.Caption = frmPld.lblBe.Tag & frmPld.cmdQMA.Caption
frmPld.lblBd.Caption = frmPld.lblBd.Tag & frmPld.cmdQMB.Caption
frmPld.lblBc.Caption = frmPld.lblBc.Tag & frmPld.cmdQMC.Caption
frmPld.lblBB.Caption = frmPld.lblBB.Tag & frmPld.cmdQMD.Caption
frmPld.lblBa.Caption = frmPld.lblBa.Tag & frmPld.cmdQME.Caption

frmPld.lblPmid.Caption = mod1.HTP.Fields("Pmid").Value
frmPld.lblGuid.Caption = mod1.HTP.Fields("Guid").Value
frmPld.txtTa.Text = mod1.HTP.Fields("BZA").Value
frmPld.txtTa.Text = mod1.HTP.Fields("BZA").Value
frmPld.txtTa.Text = mod1.HTP.Fields("BZA").Value
frmPld.txtTa.Text = mod1.HTP.Fields("BZA").Value
frmPld.txtTa.Text = mod1.HTP.Fields("BZA").Value

frmPld.lblYwy.Caption = mod1.HTP.Fields("ywy").Value
frmPld.lblUid.Caption = mod1.HTP.Fields("uid").Value
frmPld.lblLc.Caption = mod1.HTP.Fields("lc").Value
frmPld.lblLcRen.Caption = mod1.HTP.Fields("lcren").Value
frmPld.lblLcUid.Caption = mod1.HTP.Fields("lcuid").Value
frmPld.lblFwid.Caption = mod1.HTP.Fields("fwid").Value
frmPld.lblNlb.Caption = mod1.HTP.Fields("nlb").Value
frmPld.lblLcou.Caption = mod1.HTP.Fields("lcou").Value
frmPld.lblPwf.Caption = mod1.HTP.Fields("pwf").Value


tt = "select qy,bm from renyuan where username='" & frmPld.lblYwy.Caption & "' and userid='" & frmPld.lblUid.Caption & "'"
mod1.HTT.Close
mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
frmPld.lblBm.Caption = mod1.HTT.Fields("bm").Value
frmPld.lblQy.Caption = mod1.HTT.Fields("qy").Value
mod1.HTT.Close


'显示状态
If mod1.HTP.Fields("zf") = True Then
    If Lc <= 5 Then
        frmPld.lblZT.Caption = "此单在正常运作"
        frmPld.lblZT.ForeColor = &HFF0000
    ElseIf Lc = 6 Then
        frmPld.lblZT.Caption = "此单已完成"
        frmPld.lblZT.ForeColor = &HC000C0
    End If
    frmPld.lblZfyy.Visible = False
    frmPld.txtZfyy.Visible = False
    frmPld.cmdZF.Enabled = True
    frmPld.cmdSave.Enabled = True
Else
    frmPld.lblZT.Caption = "此单已经作废"
    frmPld.lblZT.ForeColor = &HFF&
    frmPld.lblZfyy.Visible = True
    frmPld.txtZfyy.Visible = True
    frmPld.cmdZF.Enabled = False
    frmPld.cmdSave.Enabled = False
End If

If mod1.HTP.Fields("xj") = False Then
    MsgBox ("这是张旧单子,请您刷新列表!")
    frmPld.Visible = False
    Exit Sub
End If

tt = "PLDBoundB(" & Pmid & ")"
frmPld.adoHp.Recordset.Close
frmPld.adoHp.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
Set frmPld.dtgSale.DataSource = frmPld.adoHp

Call modPld.PldLock '货品列表锁定

'打开采购多余表
tt = "PLDDy(" & Pmid & ")"
frmPldDy.adoDy.Recordset.Close
frmPldDy.adoDy.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
If frmPldDy.adoDy.Recordset.RecordCount > 0 Then
    Set frmPldDy.dtgDy.DataSource = frmPldDy.adoDy
    frmPld.cmdDy.Visible = True

Else
    frmPld.cmdDy.Visible = False

End If


End Sub


















Public Sub PLDQing() '配料单清空
Dim tt As String
On Error Resume Next
'frmPld.Show

frmPld.txtXmmc.Text = ""
frmPld.txtKhAdr.Text = ""
frmPld.txtHtbh.Text = ""
frmPld.txtHtze.Text = ""
frmPld.lblPmid.Caption = ""
frmPld.lblLc.Caption = "" '流程
frmPld.lblGuid.Caption = ""
frmPld.lblZT.Caption = "" '状态
frmPld.txtZfyy.Text = "" '作废原因
frmPld.txtCb.Text = "" '成本总额
frmPld.txtGdyy.Text = "" '改单原因
frmPld.lblXz.Caption = "" '合同性质
frmPld.lblKDRQ.Caption = "" '开单日期

frmPld.cmdQMA.Caption = ""
frmPld.cmdQMB.Caption = ""
frmPld.cmdQMC.Caption = ""
frmPld.cmdQMD.Caption = ""
frmPld.cmdQME.Caption = ""
frmPld.cmdYw.Caption = ""
frmPld.lblYwT.Caption = ""
frmPld.lblTa.Caption = ""
frmPld.lblTb.Caption = ""
frmPld.lblTC.Caption = ""
frmPld.lblTd.Caption = ""
frmPld.lblTe.Caption = ""
frmPld.lblBa.Caption = frmPld.lblBa.Tag
frmPld.lblBB.Caption = frmPld.lblBB.Tag
frmPld.lblBc.Caption = frmPld.lblBc.Tag
frmPld.lblBd.Caption = frmPld.lblBd.Tag
frmPld.lblBe.Caption = frmPld.lblBe.Tag
frmPld.txtTa.Text = ""
frmPld.txtTb.Text = ""
frmPld.txtTc.Text = ""
frmPld.txtTd.Text = ""
frmPld.txtTe.Text = ""
frmPld.lblPmid.Caption = ""
frmPld.lblGuid.Caption = ""
Set frmPld.dtgSale.DataSource = Nothing

'清旧单
frmPld.lblJid = ""
frmPld.cmdJa.Caption = ""
frmPld.cmdJb.Caption = ""
frmPld.cmdJc.Caption = ""
frmPld.cmdJd.Caption = ""
frmPld.cmdJe.Caption = ""
frmPld.lblJa.Caption = ""
frmPld.lblJb.Caption = ""
frmPld.lblJc.Caption = ""
frmPld.lblJd.Caption = ""
frmPld.lblJe.Caption = ""
frmPld.lblOKDRQ.Caption = ""
Set frmPld.dtgJu.DataSource = Nothing

frmPld.lblLc.Caption = ""
frmPld.lblLcRen.Caption = ""
frmPld.lblLcUid.Caption = ""
frmPld.lblFwid.Caption = ""
frmPld.lblNlb.Caption = ""
frmPld.lblLcou.Caption = "" '
frmPld.lblPwf.Caption = ""
frmPld.lblQy.Caption = ""
frmPld.lblBm.Caption = ""
frmPld.txtFcBz.Text = ""
frmPld.txtFcsl.Text = ""

End Sub

Public Sub PldLock() '货品列表锁定
Dim DD As String
Dim InHtWX As Integer
Dim InHtWB As Integer
Dim InHtLP As Integer
Dim InHtCP As Integer
On Error Resume Next
frmPld.dtgSale.Columns("产品名称").Locked = True
frmPld.dtgSale.Columns("牌号商标").Locked = True
frmPld.dtgSale.Columns("规格型号").Locked = True
frmPld.dtgSale.Columns("单位").Locked = True
frmPld.dtgSale.Columns("数量").Locked = True
frmPld.dtgSale.Columns("库存数量").Locked = True
frmPld.dtgSale.Columns("采购数量").Locked = True
frmPld.dtgSale.Columns("预计采购期").Locked = True
frmPld.dtgSale.Columns("采购到货量").Locked = True
frmPld.dtgSale.Columns("采购到货期").Locked = True
frmPld.dtgSale.Columns("供应商").Locked = True
frmPld.dtgSale.Columns("发货数量").Locked = True
frmPld.dtgSale.Columns("发货日期").Locked = True
'frmPld.dtgSale.Columns("单价").Visible = False
'frmPld.dtgSale.Columns("金额").Visible = False
frmPld.cmdSave.Enabled = False
'frmPld.dtgSale.Columns("ljmc").Locked = True
'frmPld.dtgSale.Columns("ljmc").Locked = True
frmPld.cmdAD.Visible = False
frmPld.cmdDe.Visible = False
frmPld.cmdCB.Visible = False '成本总额
frmPld.lblCb.Visible = False
frmPld.txtCb.Visible = False
frmPld.cmdGd.Visible = False
frmPld.cmdGd.Enabled = True
'''If frmPld.txtGdyy.Text <> "" Then
'''    frmPld.lblGdYY.Visible = True
'''    frmPld.txtGdyy.Visible = True
'''
'''Else
'''    frmPld.lblGdYY.Visible = False
'''    frmPld.txtGdyy.Visible = False
'''End If
DD = UCase(frmPld.txtHtbh)
InHtWX = InStr(DD, "WX")
InHtWB = InStr(DD, "WB")
InHtLP = InStr(DD, "LP")
InHtCP = InStr(DD, "CP")
'If Val(frmPld.lblLC.Caption) > 1 And Val(frmPld.lblLC.Caption) < 6 And mod1.PLA = True And (InHtWX > 0 Or InHtWB > 0) Then
'    frmPld.cmdGd.Visible = True
'End If
If Val(frmPld.lblLc.Caption) > 1 And Val(frmPld.lblLc.Caption) < 6 And mod1.PLA = True Then
    frmPld.cmdGd.Visible = True
End If
If Val(frmPld.lblLc.Caption) = 6 And mod1.PLA = True And (frmPld.lblXz.Caption = "WB" Or frmPld.lblXz.Caption = "WX") Then
    frmPld.cmdGd.Visible = True
End If

If frmPld.lblZT.Caption = "此单已经作废" Then
    frmPld.cmdSave.Enabled = False
End If
    

Select Case frmPld.lblLc.Caption
'Case 0
'    If mod1.PLA = False Then Exit Sub
'    frmPld.dtgSale.Columns("产品名称").Locked = False
'    frmPld.dtgSale.Columns("牌号商标").Locked = False
'    frmPld.dtgSale.Columns("规格型号").Locked = False
'    frmPld.dtgSale.Columns("单位").Locked = False
'    frmPld.dtgSale.Columns("数量").Locked = False
'    frmPld.cmdAD.Visible = True
'    frmPld.cmdDe.Visible = True
'    If mod1.PLA = True Then
'        frmPld.cmdSave.Enabled = True
'    End If
'Case 1
'    If mod1.PLA = False Then Exit Sub
'    frmPld.dtgSale.Columns("产品名称").Locked = False
'    frmPld.dtgSale.Columns("牌号商标").Locked = False
'    frmPld.dtgSale.Columns("规格型号").Locked = False
'    frmPld.dtgSale.Columns("单位").Locked = False
'    frmPld.dtgSale.Columns("数量").Locked = False
'    frmPld.cmdAD.Visible = True
'    frmPld.cmdDe.Visible = True
'    If mod1.PLA = True Then
'        frmPld.cmdSave.Enabled = True
'    End If
'Case 2
'    If mod1.PLA = False Then Exit Sub
'    frmPld.dtgSale.Columns("产品名称").Locked = False
'    frmPld.dtgSale.Columns("牌号商标").Locked = False
'    frmPld.dtgSale.Columns("规格型号").Locked = False
'    frmPld.dtgSale.Columns("单位").Locked = False
'    frmPld.dtgSale.Columns("数量").Locked = False
'    frmPld.cmdAD.Visible = True
'    frmPld.cmdDe.Visible = True
'    If mod1.PLA = True Then
'        frmPld.cmdSave.Enabled = True
'    End If


Case 3
    If mod1.PLB = False Then Exit Sub
    frmPld.dtgSale.Columns("库存数量").Locked = False
    If mod1.PLB = True Then
        frmPld.cmdSave.Enabled = True
    End If
Case 4
    If mod1.PLC = False Then Exit Sub
    frmPld.dtgSale.Columns("预计采购期").Locked = False
    frmPld.dtgSale.Columns("采购到货量").Locked = False
    frmPld.dtgSale.Columns("供应商").Locked = False
    frmPld.dtgSale.Columns("采购到货期").Locked = False
    
    frmPld.dtgSale.Columns("库存数量").Locked = False
    If mod1.PLC = True Then
        frmPld.cmdSave.Enabled = True
    End If

Case 5
    If mod1.PLD = False Then Exit Sub
    If mod1.PLD = True Then
        frmPld.cmdSave.Enabled = True

        frmPld.dtgSale.Columns("采购到货量").Locked = False
        frmPld.dtgSale.Columns("采购到货期").Locked = False
        frmPld.dtgSale.Columns("供应商").Locked = False
        frmPld.dtgSale.Columns("库存数量").Locked = False
        frmPld.dtgSale.Columns("领料数量").Locked = False
    End If

'Case 5
'
'        frmPld.dtgSale.Columns("发货数量").Locked = False
'        frmPld.dtgSale.Columns("发货日期").Locked = False
'    If mod1.PLE = False Then Exit Sub
'    If mod1.PLE = True Then
'
'        frmPld.cmdSave.Enabled = True
'        frmPld.cmdCB.Visible = True
'        frmPld.lblCB.Visible = True
'        frmPld.txtCB.Visible = True
''        frmPld.dtgSale.Columns("单价").Visible = True
''        frmPld.dtgSale.Columns("金额").Visible = True
'    End If
Case 6
    If mod1.PLV = False And mod1.PLE = False Then Exit Sub
        frmPld.cmdCB.Visible = True
        frmPld.lblCb.Visible = True
        frmPld.txtCb.Visible = True
End Select



End Sub












Public Sub PldJl(Pmid As Long)  '配料单保存
Dim tt As String
On Error Resume Next
tt = "PLDBoundA(" & Pmid & ")"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc

mod1.HTP.Update "xmmc", frmPld.txtXmmc.Text
mod1.HTP.Update "XMADR", frmPld.txtKhAdr.Text
mod1.HTP.Update "htze", frmPld.txtHtze.Text
mod1.HTP.Update "htbh", frmPld.txtHtbh.Text

mod1.HTP.Update "QMA", frmPld.cmdQMA.Caption '开单签字
mod1.HTP.Update "QMB", frmPld.cmdQMB.Caption
mod1.HTP.Update "QMC", frmPld.cmdQMC.Caption
mod1.HTP.Update "QMD", frmPld.cmdQMD.Caption
mod1.HTP.Update "QME", frmPld.cmdQME.Caption
mod1.HTP.Update "QMAT", frmPld.lblTa.Caption '开单签字时间
mod1.HTP.Update "QMBT", frmPld.lblTb.Caption
mod1.HTP.Update "QMCT", frmPld.lblTC.Caption
mod1.HTP.Update "QMDT", frmPld.lblTd.Caption
mod1.HTP.Update "QMET", frmPld.lblTe.Caption
mod1.HTP.Update "BZA", frmPld.txtTa.Text  '开单备注
mod1.HTP.Update "BZB", frmPld.txtTb.Text
mod1.HTP.Update "BZC", frmPld.txtTc.Text
mod1.HTP.Update "BZD", frmPld.txtTd.Text
mod1.HTP.Update "BZE", frmPld.txtTe.Text
mod1.HTP.Update "Tze", frmPld.txtCb.Text '成本总额
mod1.HTP.Update "GdYY", frmPld.txtGdyy.Text '改单原因
mod1.HTP.Update "xz", frmPld.lblXz.Caption '合同性质
mod1.HTP.Update "Pmid", frmPld.lblPmid.Caption
mod1.HTP.Update "Guid", frmPld.lblGuid.Caption
mod1.HTP.Update "LC", frmPld.lblLc.Caption '流程

mod1.HTP.UpdateBatch

'更新货品表
frmPld.adoHp.Recordset.UpdateBatch


End Sub

Public Sub PldOldBound(Pmid As Long) '旧配料单绑定
Dim tt As String
On Error Resume Next

tt = "PLDBoundA(" & Pmid & ")"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc

frmPld.cmdJa.Caption = mod1.HTP.Fields("QMA").Value
frmPld.cmdJb.Caption = mod1.HTP.Fields("QMB").Value
frmPld.cmdJc.Caption = mod1.HTP.Fields("QMC").Value
frmPld.cmdJd.Caption = mod1.HTP.Fields("QMD").Value
frmPld.cmdJe.Caption = mod1.HTP.Fields("QME").Value
frmPld.lblJa.Caption = mod1.HTP.Fields("QMAT").Value
frmPld.lblJb.Caption = mod1.HTP.Fields("QMBT").Value
frmPld.lblJc.Caption = mod1.HTP.Fields("QMCT").Value
frmPld.lblJd.Caption = mod1.HTP.Fields("QMDT").Value
frmPld.lblJe.Caption = mod1.HTP.Fields("QMET").Value
'frmPld.txtGdyy.Text = mod1.HtP.Fields("Gdyy").Value '改单原因
frmPld.lblJid.Caption = mod1.HTP.Fields("Pmid").Value
frmPld.lblOKDRQ.Caption = mod1.HTP.Fields("KRQ").Value

tt = "PLDBoundB(" & Pmid & ")"
frmPld.adoJu.Recordset.Close
frmPld.adoJu.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
Set frmPld.dtgJu.DataSource = frmPld.adoJu

End Sub
