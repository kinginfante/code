Attribute VB_Name = "modHt"
Public hTchar As String 'Sql命令
Public adoRGF As Object

Public Sub addHt() '合同评审单提交



End Sub

Public Sub htQing() '合同评审单清空
form2Htp.lblKhdh.Caption = ""
form2Htp.txtKhmc.Text = "" '客户名称
form2Htp.txtXmmc.Text = "" '项目名称
form2Htp.txtYwy.Text = "" '业务员
form2Htp.txtHtbh.Text = "" '合同编号
form2Htp.dt1.Value = mod1.HMDa
'form2Htp.comQy.Text = "" '区域
form2Htp.txtHtdate.Text = ""  '合同日期
'form2Htp.txtDdbh.Text = "" '订单编号
'form2Htp.txtDddate.Text = "" '订单日期
form2Htp.lblHtxz.Caption = "" '合同性质
form2Htp.optA(0).Value = False
form2Htp.optA(1).Value = False
form2Htp.optA(5).Value = False
'form2Htp.txtHtqy.Text = "" '合同期限起始期
'form2Htp.txtHtqy1.Text = "" '合同期限结束期
'form2Htp.comJzpb.Text = " " '机组品牌
'form2Htp.txtJzxh.Text = "" '机组型号
'form2Htp.txtJzcount.Text = "" '机组数量
'form2Htp.txtYjxh.Text = "" '压缩机型号
form2Htp.txtTian.Text = "" '交货天
form2Htp.txtJhqk.Text = "" '交货情况
form2Htp.txtMOn.Text = "" '保修期
form2Htp.txtHtze.Text = "" '合同总额
form2Htp.txtClf.Text = "" '材料费
form2Htp.txtRgf.Text = "" '人工费
form2Htp.txtCbze1.Text = "" '成本总额
form2Htp.txtCbze2.Text = "" '实际成本总额
'form2Htp.txtCbze3.Text = "" '合同成本总额
form2Htp.txtClcb1.Text = "" '材料成本
form2Htp.txtClcb2.Text = "" '实际材料成本
'form2Htp.txtClcb3.Text = "" '合同材料成本
form2Htp.txtFbje1.Text = "" '分包金额
form2Htp.txtFbje2.Text = "" '实际分包金额

'form2Htp.txtFbje3.Text = "" '合同分包金额
form2Htp.txtYf1.Text = "" '运费
form2Htp.txtYf2.Text = "" '实际运费
'form2Htp.txtYf3.Text = "" '合同运费
form2Htp.txtQt1.Text = "" '其他
form2Htp.txtQt2.Text = "" '实际其他
'form2Htp.txtQt3.Text = "" '合同其他
form2Htp.txtYj1.Text = "" '佣金
form2Htp.txtYj2.Text = "" '实际佣金
'form2Htp.txtYj3.Text = "" '合同佣金
form2Htp.txtLr1.Text = "" '项目利润
form2Htp.txtLr2.Text = "" '实际项目利润
'form2Htp.txtLr3.Text = "" '合同项目利润
form2Htp.txtTc1.Text = "" '提成
form2Htp.txtTc2.Text = "" '实际提成
'form2Htp.txtTc3.Text = "" '合同提成
form2Htp.txtJlr1.Text = ""
form2Htp.txtJlr2.Text = ""
form2Htp.txtZXF1.Text = "" '装卸费
form2Htp.txtZxF2.Text = "" '实际装卸费

form2Htp.txtCBze3.Text = ""
form2Htp.txtClcb3.Text = ""
form2Htp.txtQT3.Text = ""
form2Htp.txtZXF3.Text = ""
form2Htp.txtYf3.Text = ""
form2Htp.txtFbje3.Text = ""
form2Htp.txtTcBe.Text = 6 '提成比例
form2Htp.txtTcBe.Visible = False
form2Htp.lblTcBe.Visible = False
form2Htp.UpDa.Visible = False
form2Htp.txtTc1.Visible = True
form2Htp.txtTc2.Visible = True

'开票类型
form2Htp.optLa.Value = False
form2Htp.optLb.Value = False
form2Htp.optLc.Value = False
form2Htp.optLD.Value = False
form2Htp.optLE.Value = False
'form2Htp.frmKP.Visible = False
form2Htp.chkDzf.Value = False

form2Htp.chkA.Caption = "" '业务员签名
form2Htp.chkA.Tag = "" '现在业务员
form2Htp.chkB.Caption = "" '销售经理签名
form2Htp.chkB.Tag = "" '现在销售经理
form2Htp.chkC.Caption = "" '商务经理签名
form2Htp.chkD.Caption = "" '总经理签名
form2Htp.chkE.Caption = "" '技术支持签字
form2Htp.lblYw.Caption = ""
form2Htp.lblJZ.Caption = ""
form2Htp.lblJl.Caption = ""
form2Htp.lblYZ.Caption = ""
form2Htp.lblZJ.Caption = ""

'添加字段
form2Htp.txtAdd1.Text = ""
form2Htp.txtAz1.Text = ""
form2Htp.txtAdd2.Text = ""
form2Htp.txtAz2.Text = ""
form2Htp.txtAdd3.Text = ""
form2Htp.txtAz3.Text = ""
form2Htp.txtAdd4.Text = ""
form2Htp.txtAz4.Text = ""
form2Htp.txtAdd5.Text = ""
form2Htp.txtAz5.Text = ""
form2Htp.txtFkBz.Text = "" '付款条件备注
form2Htp.lblHid.Caption = ""
form2Htp.txtTcRQ.Text = "提成取现日期"

form2Htp.optP.BackColor = &H8000000F
End Sub


Public Sub htBound() '合同评审单字段绑定(直接赋值，不采用绑定方法)
On Error Resume Next
Call mod1.zhuDa(1, mod1.HTP.Fields("htbh").Value)
form2Htp.txtKhmc.Text = mod1.HTP.Fields(0).Value '客户名称
form2Htp.lblKhdh.Caption = mod1.HTP.Fields("khdh").Value '客户代号
form2Htp.txtXmmc.Text = mod1.HTP.Fields("xmmc").Value '项目名称
form2Htp.txtYwy.Text = mod1.HTP.Fields("YwY").Value '业务员
form2Htp.txtHtbh.Text = mod1.HTP.Fields(2).Value '合同编号

form2Htp.txtHtdate.Text = Format(mod1.HTP.Fields(3).Value, "Long Date") '合同日期
form2Htp.dt1.Value = mod1.HTP.Fields(3).Value
'form2Htp.txtDdbh.Text = mod1.HtP.Fields(4).Value '订单编号
'form2Htp.txtDddate.Text = Format(mod1.HtP.Fields(5).Value, "Long Date") '订单日期
'form2Htp.dt2.Value = mod1.HtP.Fields(5).Value

form2Htp.lblHtxz.Caption = mod1.HTP.Fields(6).Value '合同性质
'Select Case Mid(mod1.HtP.Fields(2).Value, 4, 2)
'Case "HS"
'form2Htp.opta(0).Value = True
'Case "GC"
'form2Htp.opta(1).Value = True
'Case "FB"
'form2Htp.opta(2).Value = True
'Case "WB"
'form2Htp.opta(3).Value = True
'Case "WX"
'form2Htp.opta(4).Value = True
'End Select
Select Case form2Htp.lblHtxz.Caption
Case "A. 零配件合同"
form2Htp.optA(0).Value = True
Case "B1.工程合同"
form2Htp.optA(1).Value = True
'Case "B2.分包合同"
'form2Htp.optA(2).Value = True
Case "C. 维保合同"
form2Htp.optA(3).Value = True
Case "D. 维修合同"
form2Htp.optA(4).Value = True
Case "E. 产品合同"
form2Htp.optA(5).Value = True
End Select

form2Htp.comQy.Text = mod1.HTP.Fields("qy").Value '区域

'form2Htp.txtHtqy.Text = Format(mod1.HtP.Fields(7).Value, "Long Date") '合同期限起始期
'form2Htp.dt3.Value = mod1.HtP.Fields(7).Value
'form2Htp.txtHtqy1.Text = Format(mod1.HtP.Fields(8).Value, "Long Date") '合同期限结束期
'form2Htp.dt4.Value = mod1.HtP.Fields(8).Value
'form2Htp.comJzpb.Text = mod1.HtP.Fields(9).Value '机组品牌
'form2Htp.txtJzxh.Text = mod1.HtP.Fields(10).Value '机组型号
'form2Htp.txtJzcount.Text = mod1.HtP.Fields(11).Value '机组数量
'form2Htp.txtYjxh.Text = mod1.HtP.Fields(12).Value '压缩机型号
form2Htp.txtTian.Text = mod1.HTP.Fields(13).Value '交货天
form2Htp.txtJhqk.Text = mod1.HTP.Fields(14).Value '交货情况
form2Htp.txtMOn.Text = mod1.HTP.Fields(15).Value '保修期
If mod1.HTP.Fields(16).Value = 0 Then
form2Htp.txtHtze.Text = ""
Else
form2Htp.txtHtze.Text = mod1.HTP.Fields(16).Value '合同总额
End If

frmFuK.lblHtze.Caption = Round(mod1.HTP.Fields(16).Value, 2)

If mod1.HTP.Fields(17).Value = 0 Then
form2Htp.txtClf.Text = ""
Else
form2Htp.txtClf.Text = mod1.HTP.Fields(17).Value '材料费
End If

If mod1.HTP.Fields(18).Value = 0 Then
form2Htp.txtRgf.Text = ""
Else
form2Htp.txtRgf.Text = mod1.HTP.Fields(18).Value '人工费
End If

If mod1.HTP.Fields(19).Value = 0 Then
form2Htp.txtCbze1.Text = ""
Else
form2Htp.txtCbze1.Text = mod1.HTP.Fields(19).Value '成本总额
End If

If mod1.HTP.Fields(20).Value = 0 Then
form2Htp.txtCbze2.Text = ""
Else
form2Htp.txtCbze2.Text = mod1.HTP.Fields("cbze1").Value '实际成本总额
End If

If mod1.HTP.Fields(52).Value = 0 Then
'form2Htp.txtCbze3.Text = ""
Else
'form2Htp.txtCbze3.Text = mod1.HtP.Fields(52).Value '合同成本总额
End If

If mod1.HTP.Fields(21).Value = 0 Then
form2Htp.txtClcb1.Text = ""
Else
form2Htp.txtClcb1.Text = mod1.HTP.Fields(21).Value '材料成本
End If

If mod1.HTP.Fields(22).Value = 0 Then
form2Htp.txtClcb2.Text = ""
Else
form2Htp.txtClcb2.Text = mod1.HTP.Fields(22).Value '实际材料成本
End If

If mod1.HTP.Fields(53).Value = 0 Then
'form2Htp.txtClcb3.Text = ""
Else
'form2Htp.txtClcb3.Text = mod1.HtP.Fields(53).Value '合同材料成本
End If

If mod1.HTP.Fields(23).Value = 0 Then
form2Htp.txtFbje1.Text = ""
Else
form2Htp.txtFbje1.Text = mod1.HTP.Fields(23).Value '分包金额
End If

If mod1.HTP.Fields(24).Value = 0 Then
form2Htp.txtFbje2.Text = ""
Else
form2Htp.txtFbje2.Text = mod1.HTP.Fields(24).Value '实际分包金额
End If

If mod1.HTP.Fields(54).Value = 0 Then
'form2Htp.txtFbje3.Text = ""
Else
'form2Htp.txtFbje3.Text = mod1.HtP.Fields(54).Value '合同分包金额
End If


If mod1.HTP.Fields(25).Value = 0 Then
form2Htp.txtYf1.Text = ""
Else
form2Htp.txtYf1.Text = mod1.HTP.Fields(25).Value '运费
End If

If mod1.HTP.Fields(26).Value = 0 Then
form2Htp.txtYf2.Text = ""
Else
form2Htp.txtYf2.Text = mod1.HTP.Fields(26).Value '实际运费
End If

If mod1.HTP.Fields(55).Value = 0 Then
'form2Htp.txtYf3.Text = ""
Else
'form2Htp.txtYf3.Text = mod1.HtP.Fields(55).Value '合同运费
End If

If mod1.HTP.Fields(27).Value = 0 Then
form2Htp.txtQt1.Text = ""
Else
form2Htp.txtQt1.Text = mod1.HTP.Fields(27).Value '其他
End If

If mod1.HTP.Fields(28).Value = 0 Then
form2Htp.txtQt2.Text = ""
Else
form2Htp.txtQt2.Text = mod1.HTP.Fields(28).Value '实际其他
End If

If mod1.HTP.Fields(56).Value = 0 Then
'form2Htp.txtQt3.Text = ""
Else
'form2Htp.txtQt3.Text = mod1.HtP.Fields(56).Value '合同其他
End If

If mod1.HTP.Fields(29).Value = 0 Then
form2Htp.txtYj1.Text = ""
Else
form2Htp.txtYj1.Text = mod1.HTP.Fields(29).Value '佣金
End If

If mod1.HTP.Fields(30).Value = 0 Then
form2Htp.txtYj2.Text = ""
Else
form2Htp.txtYj2.Text = mod1.HTP.Fields(30).Value '实际佣金
End If

If mod1.HTP.Fields(57).Value = 0 Then
'form2Htp.txtYj3.Text = ""
Else
'form2Htp.txtYj3.Text = mod1.HtP.Fields(57).Value '合同佣金
End If

If mod1.HTP.Fields(31).Value = 0 Then
form2Htp.txtLr1.Text = ""
Else
form2Htp.txtLr1.Text = mod1.HTP.Fields(31).Value '项目利润
End If

If mod1.HTP.Fields(32).Value = 0 Then
form2Htp.txtLr2.Text = ""
Else
form2Htp.txtLr2.Text = mod1.HTP.Fields(32).Value '实际项目利润
End If

If mod1.HTP.Fields(58).Value = 0 Then
'form2Htp.txtLr3.Text = ""
Else
'form2Htp.txtLr3.Text = mod1.HtP.Fields(58).Value '合同项目利润
End If

form2Htp.txtJlr1.Text = mod1.HTP.Fields("jlr1").Value
form2Htp.txtJlr2.Text = mod1.HTP.Fields("jlr2").Value



form2Htp.txtCBze3.Text = mod1.HTP.Fields("cbze3").Value
form2Htp.txtClcb3.Text = mod1.HTP.Fields("clcb3").Value
form2Htp.txtQT3.Text = mod1.HTP.Fields("qt3").Value
form2Htp.txtZXF3.Text = mod1.HTP.Fields("zxf3").Value
form2Htp.txtYf3.Text = mod1.HTP.Fields("yf3").Value
form2Htp.txtFbje3.Text = mod1.HTP.Fields("fbje3").Value
form2Htp.txtTcBe.Text = mod1.HTP.Fields("tcbe").Value '提成比例

If mod1.HTP.Fields(33).Value = 0 Then
form2Htp.txtTc1.Text = ""
Else
form2Htp.txtTc1.Text = mod1.HTP.Fields(33).Value '提成
End If

If mod1.HTP.Fields(34).Value = 0 Then
form2Htp.txtTc2.Text = ""
Else
form2Htp.txtTc2.Text = mod1.HTP.Fields(34).Value '实际提成
End If

If mod1.HTP.Fields(59).Value = 0 Then
'form2Htp.txtTc3.Text = ""
Else
'form2Htp.txtTc3.Text = mod1.HtP.Fields(59).Value '合同提成
End If

form2Htp.txtZXF1.Text = mod1.HTP.Fields("rgF").Value '预计装卸费
form2Htp.txtZxF2.Text = mod1.HTP.Fields("rgF1").Value '实际装卸费

'开票类型
If mod1.HTP.Fields("fpLX").Value = "增值发票" Then
form2Htp.optLa.Value = True
ElseIf mod1.HTP.Fields("fpLX").Value = "商业发票" Then
form2Htp.optLb.Value = True
ElseIf mod1.HTP.Fields("fpLX").Value = "服务发票" Then
form2Htp.optLc.Value = True
ElseIf mod1.HTP.Fields("fpLX").Value = "其它" Then
form2Htp.optLD.Value = True
ElseIf mod1.HTP.Fields("fpLX").Value = "不开票" Then
form2Htp.optLE.Value = True
End If
'末开票到帐
If mod1.HTP.Fields("dzF").Value = 1 Then
    form2Htp.chkDzf.Value = 1
ElseIf mod1.HTP.Fields("dzF").Value = 0 Then
    form2Htp.chkDzf.Value = 0
End If

If mod1.HTP.Fields(35).Value <> "" Then '业务员签字
form2Htp.chkA.Caption = mod1.HTP.Fields(35).Value
form2Htp.chkA.Tag = mod1.HTP.Fields("xywy").Value
form2Htp.chkA.Value = 1
Else
form2Htp.chkA.Caption = ""
form2Htp.chkA.Value = 0
End If

If mod1.HTP.Fields("JzQz").Value <> "" Then '技术支持签字
form2Htp.chkE.Caption = mod1.HTP.Fields("JzQz").Value
form2Htp.chkE.Value = 1
Else
form2Htp.chkE.Caption = ""
form2Htp.chkE.Value = 0
End If

If mod1.HTP.Fields(36).Value <> "" Then '销售经理签字
form2Htp.chkB.Caption = mod1.HTP.Fields(36).Value
form2Htp.chkB.Tag = mod1.HTP.Fields("xjlq").Value
form2Htp.chkB.Value = 1
Else
form2Htp.chkB.Caption = ""
form2Htp.chkB.Value = 0
End If
If mod1.HTP.Fields(37).Value <> "" Then '商务经理签字
form2Htp.chkC.Caption = mod1.HTP.Fields(37).Value
form2Htp.chkC.Value = 1
Else
form2Htp.chkC.Caption = ""
form2Htp.chkC.Value = 0
End If
If mod1.HTP.Fields(38).Value <> "" Then '总经理签字
form2Htp.chkD.Caption = mod1.HTP.Fields(38).Value
form2Htp.chkD.Value = 1
Else
form2Htp.chkD.Caption = ""
form2Htp.chkD.Value = 0
End If

form2Htp.lblYw.Caption = mod1.HTP.Fields("ywDa").Value '销售签字日期
form2Htp.lblJZ.Caption = mod1.HTP.Fields("JzDa").Value '技术支持签字日期
form2Htp.lblJl.Caption = mod1.HTP.Fields("JlDa").Value '销售经理签字日期
form2Htp.lblYZ.Caption = mod1.HTP.Fields("YzDa").Value '商务经理签字日期
form2Htp.lblZJ.Caption = mod1.HTP.Fields("ZjDa").Value '总经理签字日期

'form2Htp.txtYwy.Text = form2Htp.chkA.Caption  '业务员

'添加的字段
If mod1.HTP.Fields(39).Value = 0 Then
form2Htp.txtAdd1.Text = ""
Else
form2Htp.txtAdd1.Text = mod1.HTP.Fields(39).Value
End If

If mod1.HTP.Fields(40).Value = 0 Then
form2Htp.txtAz1.Text = ""
Else
form2Htp.txtAz1.Text = mod1.HTP.Fields(40).Value
End If

If mod1.HTP.Fields(41).Value = 0 Then
form2Htp.txtAdd2.Text = ""
Else
form2Htp.txtAdd2.Text = mod1.HTP.Fields(41).Value
End If

If mod1.HTP.Fields(42).Value = 0 Then
form2Htp.txtAz2.Text = ""
Else
form2Htp.txtAz2.Text = mod1.HTP.Fields(42).Value
End If

If mod1.HTP.Fields(43).Value = 0 Then
form2Htp.txtAdd3.Text = ""
Else
form2Htp.txtAdd3.Text = mod1.HTP.Fields(43).Value
End If

If mod1.HTP.Fields(44).Value = 0 Then
form2Htp.txtAz3.Text = ""
Else
form2Htp.txtAz3.Text = mod1.HTP.Fields(44).Value
End If

If mod1.HTP.Fields(45).Value = 0 Then
form2Htp.txtAdd4.Text = ""
Else
form2Htp.txtAdd4.Text = mod1.HTP.Fields(45).Value
End If

If mod1.HTP.Fields(46).Value = 0 Then
form2Htp.txtAz4.Text = ""
Else
form2Htp.txtAz4.Text = mod1.HTP.Fields(46).Value
End If

If mod1.HTP.Fields(47).Value = 0 Then
form2Htp.txtAdd5.Text = ""
Else
form2Htp.txtAdd5.Text = mod1.HTP.Fields(47).Value
End If

If mod1.HTP.Fields(48).Value = 0 Then
form2Htp.txtAz5.Text = ""
Else
form2Htp.txtAz5.Text = mod1.HTP.Fields(48).Value
End If

form2Htp.txtFkBz.Text = mod1.HTP.Fields("fkBz").Value '付款条件备注
form2Htp.comQy.Text = mod1.HTP.Fields("qy").Value '区域
form2Htp.lblBM.Caption = mod1.HTP.Fields("bm").Value

If IsNull(mod1.HTP.Fields("TCRQ").Value) = False Then
    form2Htp.txtTcRQ.Text = mod1.HTP.Fields("TCRQ").Value '提成取现日期
End If

If mod1.HTP.Fields("jTf").Value = True Then
    form2Htp.cmdCount.Caption = "已结算"
    form2Htp.cmdCount.Enabled = False
Else
    form2Htp.cmdCount.Caption = "计算"
    form2Htp.cmdCount.Enabled = True
End If

'如果为旧合同,则评审阶段设为蓝色
If mod1.HTP.Fields("XGG").Value = True Then
    form2Htp.optP.BackColor = &HC0FFFF
ElseIf mod1.HTP.Fields("XGG").Value = 0 Then
    form2Htp.optP.BackColor = &H8000000F
End If

form2Htp.optP.Enabled = False
form2Htp.optG.Enabled = False
form2Htp.optZ.Enabled = False
form2Htp.optW.Enabled = False

'合同执行否
If mod1.HTP.Fields(51).Value = 0 Then
    form2Htp.optP.Value = True
'    form2Htp.optP.Enabled = True
'    '如果都签字了,而且为盖章者打开,则可以进行盖章
'    If (form2Htp.chkC.Caption <> "" And form2Htp.txtHtze.Text < 10000 Or form2Htp.chkD.Caption <> "") And _
'        mod1.KGZ = True Then
'        form2Htp.optG.Enabled = True
'    End If
    
ElseIf mod1.HTP.Fields("htF").Value = 9 Then
    form2Htp.optG.Value = True
'    form2Htp.optG.Enabled = True
'    If mod1.KZX = True Then
'        form2Htp.optZ.Enabled = True
'    End If

ElseIf mod1.HTP.Fields(51).Value = 1 Then
    form2Htp.optZ.Value = True
'    form2Htp.optZ.Enabled = True
'    If mod1.KWC = True Then
'        form2Htp.optW.Enabled = True
'    End If

ElseIf mod1.HTP.Fields(51).Value = 2 Then
    form2Htp.optW.Value = True
    form2Htp.optW.Enabled = True
End If


form2Htp.frmZt.Enabled = True
''如果合同评审没有全通过或者不是小吴，则“合同执行否”字段不能编辑
'If form2Htp.chkD.Caption <> "" And frmLogin.Combo1.Text = "胡颖" Then
'form2Htp.frmZt.Enabled = True
'Else
'末开票到帐否
If mod1.HTP.Fields("dzF").Value = True Then
    form2Htp.chkDzf.Value = 1
ElseIf mod1.HTP.Fields("dzF").Value = False Then
    form2Htp.chkDzf.Value = 0
End If
End Sub























Public Sub lianJ() '判断收款表中每条记录是否与应收表中一一对应，如果否，则删除

End Sub




Public Sub HtF() '判断htping,htping1,yiFk,htSale表中的htF,是否一致，如果否，则将全部对应htping的值


End Sub























Public Sub qianKuan() '更新欠款
Dim ladate As Date
Dim cadate As Integer
Dim LT As String
On Error Resume Next
'如果已收款，则确定最后付款日期,再进行欠款统计
If frmFuK.adoYf.Recordset.Fields(5).Value = True Then
    frmFuK.adoYf.Recordset.MoveLast
    ladate = frmFuK.adoYf.Recordset.Fields(0).Value
        Do While Not frmFuK.adoYf.Recordset.BOF
        cadate = DateDiff("D", frmFuK.adoYf.Recordset.Fields(3).Value, ladate)
        frmFuK.adoYf.Recordset.Update "laRq", ladate
        '进行欠款计算
            If cadate <= 0 Then
            frmFuK.adoYf.Recordset.Update "qianKuan1", 1
            frmFuK.adoYf.Recordset.Update "qianKuan2", 1
            frmFuK.adoYf.Recordset.Update "qianKuan3", 1
            frmFuK.adoYf.Recordset.Update "qianKuan4", 1
            frmFuK.adoYf.Recordset.Update "qianKuan5", 1
            frmFuK.adoYf.Recordset.Update "qianKuan6", 1
            ElseIf cadate > 0 And cadate <= 30 Then
            frmFuK.adoYf.Recordset.Update "qianKuan1", 0
            frmFuK.adoYf.Recordset.Update "qianKuan2", 1
            frmFuK.adoYf.Recordset.Update "qianKuan3", 1
            frmFuK.adoYf.Recordset.Update "qianKuan4", 1
            frmFuK.adoYf.Recordset.Update "qianKuan5", 1
            frmFuK.adoYf.Recordset.Update "qianKuan6", 1
            ElseIf cadate > 30 And cadate <= 91 Then
            frmFuK.adoYf.Recordset.Update "qianKuan1", 0
            frmFuK.adoYf.Recordset.Update "qianKuan2", 0
            frmFuK.adoYf.Recordset.Update "qianKuan3", 1
            frmFuK.adoYf.Recordset.Update "qianKuan4", 1
            frmFuK.adoYf.Recordset.Update "qianKuan5", 1
            frmFuK.adoYf.Recordset.Update "qianKuan6", 1
            ElseIf cadate > 91 And cadate <= 183 Then
            frmFuK.adoYf.Recordset.Update "qianKuan1", 0
            frmFuK.adoYf.Recordset.Update "qianKuan2", 0
            frmFuK.adoYf.Recordset.Update "qianKuan3", 0
            frmFuK.adoYf.Recordset.Update "qianKuan4", 1
            frmFuK.adoYf.Recordset.Update "qianKuan5", 1
            frmFuK.adoYf.Recordset.Update "qianKuan6", 1
            ElseIf cadate > 183 And cadate <= 365 Then
            frmFuK.adoYf.Recordset.Update "qianKuan1", 0
            frmFuK.adoYf.Recordset.Update "qianKuan2", 0
            frmFuK.adoYf.Recordset.Update "qianKuan3", 0
            frmFuK.adoYf.Recordset.Update "qianKuan4", 0
            frmFuK.adoYf.Recordset.Update "qianKuan5", 1
            frmFuK.adoYf.Recordset.Update "qianKuan6", 1
            ElseIf cadate > 365 And cadate <= 730 Then
            frmFuK.adoYf.Recordset.Update "qianKuan1", 0
            frmFuK.adoYf.Recordset.Update "qianKuan2", 0
            frmFuK.adoYf.Recordset.Update "qianKuan3", 0
            frmFuK.adoYf.Recordset.Update "qianKuan4", 0
            frmFuK.adoYf.Recordset.Update "qianKuan5", 0
            frmFuK.adoYf.Recordset.Update "qianKuan6", 1
            ElseIf cadate > 730 Then
            frmFuK.adoYf.Recordset.Update "qianKuan1", 0
            frmFuK.adoYf.Recordset.Update "qianKuan2", 0
            frmFuK.adoYf.Recordset.Update "qianKuan3", 0
            frmFuK.adoYf.Recordset.Update "qianKuan4", 0
            frmFuK.adoYf.Recordset.Update "qianKuan5", 0
            frmFuK.adoYf.Recordset.Update "qianKuan6", 0
            End If
        
        frmFuK.adoYf.Recordset.UpdateBatch
        frmFuK.adoYf.Recordset.MovePrevious
        Loop
        
'如果款未收，则更据当前日期来确定欠款
ElseIf frmFuK.adoYf.Recordset.Fields(5).Value = False Then
    frmFuK.adoYf.Recordset.MoveFirst
     
    Do While Not frmFuK.adoYf.Recordset.EOF
    cadate = DateDiff("D", frmFuK.adoYf.Recordset.Fields(3).Value, Date)
     frmFuK.adoYf.Recordset.Update "qianKuan1", 0
    frmFuK.adoYf.Recordset.Update "qianKuan2", 0
    frmFuK.adoYf.Recordset.Update "qianKuan3", 0
    frmFuK.adoYf.Recordset.Update "qianKuan4", 0
    frmFuK.adoYf.Recordset.Update "qianKuan5", 0
    frmFuK.adoYf.Recordset.Update "qianKuan6", 0
    frmFuK.adoYf.Recordset.UpdateBatch
    If frmFuK.adoYf.Recordset.Fields(0).Value = #12/1/2004# Then
    kk = "aaa"
    End If
    If cadate < 0 Then
            frmFuK.adoYf.Recordset.Update "qianKuan1", 1
           frmFuK.adoYf.Recordset.Update "qianKuan2", 1
           frmFuK.adoYf.Recordset.Update "qianKuan3", 1
           frmFuK.adoYf.Recordset.Update "qianKuan4", 1
           frmFuK.adoYf.Recordset.Update "qianKuan5", 1
           frmFuK.adoYf.Recordset.Update "qianKuan6", 1
           
    ElseIf cadate >= 0 And cadate < 30 Then
            frmFuK.adoYf.Recordset.Update "qianKuan2", 1
           frmFuK.adoYf.Recordset.Update "qianKuan3", 1
           frmFuK.adoYf.Recordset.Update "qianKuan4", 1
           frmFuK.adoYf.Recordset.Update "qianKuan5", 1
           frmFuK.adoYf.Recordset.Update "qianKuan6", 1
    ElseIf cadate >= 30 And cadate < 91 Then
            frmFuK.adoYf.Recordset.Update "qianKuan3", 1
           frmFuK.adoYf.Recordset.Update "qianKuan4", 1
           frmFuK.adoYf.Recordset.Update "qianKuan5", 1
           frmFuK.adoYf.Recordset.Update "qianKuan6", 1
    ElseIf cadate >= 91 And cadate < 183 Then
            frmFuK.adoYf.Recordset.Update "qianKuan4", 1
           frmFuK.adoYf.Recordset.Update "qianKuan5", 1
           frmFuK.adoYf.Recordset.Update "qianKuan6", 1
    ElseIf cadate >= 183 And cadate < 365 Then
            frmFuK.adoYf.Recordset.Update "qianKuan5", 1
           frmFuK.adoYf.Recordset.Update "qianKuan6", 1
    ElseIf cadate >= 365 And cadate < 730 Then
            frmFuK.adoYf.Recordset.Update "qianKuan6", 1
    End If
    frmFuK.adoYf.Recordset.UpdateBatch
    frmFuK.adoYf.Recordset.MoveNext
    Loop
End If

'更新资金流量表的欠款
frmFuK.adoYf.Recordset.MoveFirst
Do While Not frmFuK.adoYf.Recordset.EOF
LT = "update llb1 set qianKuan1=1 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'" '当天
If frmFuK.adoYf.Recordset.Fields(12).Value = 0 Then
LT = ""
LT = "update llb1 set qianKuan1=0 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
LT = ""
LT = "update llb1 set qianKuan2=1 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'" '当月
If frmFuK.adoYf.Recordset.Fields(13).Value = 0 Then
LT = ""
LT = "update llb1 set qianKuan2=0 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
LT = ""
LT = "update llb1 set qianKuan3=1 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'" '3月
If frmFuK.adoYf.Recordset.Fields(14).Value = 0 Then
LT = ""
LT = "update llb1 set qianKuan3=0 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
LT = ""
LT = "update llb1 set qianKuan4=1 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'" '半年
If frmFuK.adoYf.Recordset.Fields(15).Value = 0 Then
LT = ""
LT = "update llb1 set qianKuan4=0 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
LT = ""
LT = "update llb1 set qianKuan5=1 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'" '1年
If frmFuK.adoYf.Recordset.Fields(16).Value = 0 Then
LT = ""
LT = "update llb1 set qianKuan5=0 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
LT = ""
LT = "update llb1 set qianKuan6=1 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'" '2年
If frmFuK.adoYf.Recordset.Fields(17).Value = 0 Then
LT = ""
LT = "update llb1 set qianKuan6=0 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

frmFuK.adoYf.Recordset.MoveNext
Loop
End Sub

Public Sub tgYik(daDate As Variant)    '统计当前应收

End Sub


Public Sub gxAdd() '销售合同添加数据


End Sub
















































Public Sub gxQing() '销售合同字段清空
Dim oo As Integer
htgX.txtGF.Text = "" '供方
htgX.comKhmc.Text = "" '客户名称
If htgX.comKhmc.ListCount > 0 Then
    For oo = htgX.comKhmc.ListCount - 1 To 0 Step -1
        htgX.comKhmc.RemoveItem oo
    Next
End If
htgX.comKhmc.Locked = False
htgX.txtHtbh.Text = "" '合同编号
htgX.txtQyDD.Text = "" '签约地点
htgX.txtXF.Text = "" '需方
'htgX.DTPQdDate.Value = "" '签订时间
htgX.txtHg.Text = "" '合计
htgX.lblDx.Caption = "" '合计大写
htgX.txtT2.Text = "" '二、质量要求技术标准
htgX.txtZBQ.Text = "" '保质期
htgX.txtT3.Text = "" '三、供方对质量负责的条件和期限
'htgX.txtT4.Text = "" '四、交(提)货方式
htgX.txtT5.Text = "" '五、运输方式及到达站（港）的费用负担
htgX.txtT6.Text = "" '六、合理损耗计算方法
htgX.txtT7.Text = "" '七、包装标准、包装物的供应与回收和费用负担
htgX.txtT8.Text = "" '八、验收方式及提出异议期限
htgX.txtT9.Text = "" '九、随机备品、配件工具数量及供应办法
htgX.txtT10.Text = "" '十、结算方式及期限
htgX.txtT11.Text = "" '十一、如需提供担保，另立合同担保书，作为本合同附件
htgX.txtT12.Text = "" '十二、违约责任
htgX.txtT13.Text = "" '十三、解决合同纠纷的方式
htgX.txtT14.Text = "" '十四、其它约定事项
htgX.txtGdwMc.Text = "" '单位名称
htgX.txtGdwAdr.Text = "" '单位地址
htgX.txtGfdBr.Text = "" '法定代表人
htgX.txtGdw.Text = "" '电话
htgX.txtGFX.Text = "" '传真
htgX.txtGFH.Text = "" '国税号
htgX.txtGkhYY.Text = "" '开户银行
htgX.txtGZH.Text = "" '账号
htgX.txtGyzBM.Text = "" '邮政编码
htgX.txtGwiTo.Text = "" '委托代理人
htgX.txtXdwMc.Text = ""
htgX.txtXdwAdr.Text = ""
htgX.txtXGfdBr.Text = ""
htgX.txtXGdW.Text = ""
htgX.txtXGFX.Text = ""
htgX.txtXGFH.Text = ""
htgX.txtXGkhYY.Text = ""
htgX.txtXGZH.Text = ""
htgX.txtXGyzBM.Text = ""
htgX.lblKhdh.Caption = ""
htgX.txtXGwiTo.Text = ""
'htgX.dtpYXQ.Value = "" '有效期

End Sub

























Public Sub gxBound() '销售合同字段绑定
Dim tt As String
On Error Resume Next
tt = "Select * from gxHt where htBh='" & form2Htp.txtHtbh.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
htgX.txtGF.Text = mod1.HTP.Fields("GF").Value '供方
htgX.txtHtbh.Text = mod1.HTP.Fields("htBh").Value '合同编号
htgX.txtQyDD.Text = mod1.HTP.Fields("qyDD").Value '签约地点
htgX.txtXF.Text = mod1.HTP.Fields("XF").Value '需方
htgX.comKhmc.Text = mod1.HTP.Fields("khmc").Value '客户名称
htgX.DTPQdDate.Value = mod1.HTP.Fields("qdDate").Value '签订时间
htgX.txtHg.Text = mod1.HTP.Fields("HG").Value '合计
htgX.lblDx.Caption = mod1.HTP.Fields("DHG").Value '合计大写
htgX.txtT2.Text = mod1.HTP.Fields("T2").Value '二、质量要求技术标准
htgX.txtZBQ.Text = mod1.HTP.Fields("ZBQ").Value '保质期
htgX.txtT3.Text = mod1.HTP.Fields("T3").Value '三、供方对质量负责的条件和期限
htgX.txtT4.Text = mod1.HTP.Fields("T4").Value '四、交(提)货方式
htgX.txtT5.Text = mod1.HTP.Fields("T5").Value '五、运输方式及到达站（港）的费用负担
htgX.txtT6.Text = mod1.HTP.Fields("T6").Value '六、合理损耗计算方法
htgX.txtT7.Text = mod1.HTP.Fields("T7").Value '七、包装标准、包装物的供应与回收和费用负担
htgX.txtT8.Text = mod1.HTP.Fields("T8").Value '八、验收方式及提出异议期限
htgX.txtT9.Text = mod1.HTP.Fields("T9").Value '九、随机备品、配件工具数量及供应办法
htgX.txtT10.Text = mod1.HTP.Fields("T10").Value '十、结算方式及期限
htgX.txtT11.Text = mod1.HTP.Fields("T11").Value '十一、如需提供担保，另立合同担保书，作为本合同附件
htgX.txtT12.Text = mod1.HTP.Fields("T12").Value '十二、违约责任
htgX.txtT13.Text = mod1.HTP.Fields("T13").Value '十三、解决合同纠纷的方式
htgX.txtT14.Text = mod1.HTP.Fields("T14").Value '十四、其它约定事项
htgX.txtGdwMc.Text = mod1.HTP.Fields("GdwMc").Value  '单位名称
htgX.txtGdwAdr.Text = mod1.HTP.Fields("GdwAdr").Value '单位地址
htgX.txtGfdBr.Text = mod1.HTP.Fields("GfdBr").Value '法定代表人
htgX.txtGdw.Text = mod1.HTP.Fields("GdW").Value '电话
htgX.txtGFX.Text = mod1.HTP.Fields("GFX").Value '传真
htgX.txtGFH.Text = mod1.HTP.Fields("GFH").Value '国税号
htgX.txtGkhYY.Text = mod1.HTP.Fields("GkhYY").Value '开户银行
htgX.txtGZH.Text = mod1.HTP.Fields("GZH").Value '账号
htgX.txtGyzBM.Text = mod1.HTP.Fields("GyzBM").Value '邮政编码
htgX.txtGwiTo.Text = mod1.HTP.Fields("GwiTo").Value '委托代理人
htgX.txtXdwMc.Text = mod1.HTP.Fields("XdwMc").Value
htgX.txtXdwAdr.Text = mod1.HTP.Fields("XdwAdr").Value
htgX.txtXGfdBr.Text = mod1.HTP.Fields("XGfdBr").Value
htgX.txtXGdW.Text = mod1.HTP.Fields("XGdW").Value
htgX.txtXGFX.Text = mod1.HTP.Fields("XGFX").Value
htgX.txtXGFH.Text = mod1.HTP.Fields("XGFH").Value
htgX.txtXGkhYY.Text = mod1.HTP.Fields("XGkhYY").Value
htgX.txtXGZH.Text = mod1.HTP.Fields("XGZH").Value
htgX.txtXGyzBM.Text = mod1.HTP.Fields("XGyzBM").Value
htgX.txtXGwiTo.Text = mod1.HTP.Fields("XGwiTo").Value
htgX.dtpYXQ.Value = mod1.HTP.Fields("YXQ").Value '有效期
htgX.lblKhdh.Caption = form2Htp.txtHtbh.Text
'If Len(htgX.lblKhdh.Caption) < 5 Then
'    MsgBox ("该合同评审单有问题,请与马晓聪联系!")
'End If
'更新产品表
Set htgX.dtgSale.DataSource = form2Htp.adoSale

End Sub
















Public Sub wbQing() '维保合同字段清空
On Error Resume Next

wbHTP.txtKhmc.Text = "" '客户名称
wbHTP.txtXmmc.Text = "" '项目名称
wbHTP.txtKhdm.Text = "" '客户代码


'合同签订日期
wbHTP.txtHtdate.Text = ""
wbHTP.dt1.Value = mod1.HMDa
'维保期限
wbHTP.dt3.Value = mod1.HMDa
wbHTP.dt4.Value = mod1.HMDa
wbHTP.txtGLG.Text = "" '管理公司
wbHTP.txtMOn.Text = "" '维修质保期
wbHTP.txtADR.Text = "" '项目地址
wbHTP.txtHtze.Text = "" '合同总金额
wbMx.txtFkBz.Text = "" '付款条件
wbHTP.txtCbze1.Text = "" '成本总额
wbHTP.txtCbze2.Text = ""
wbHTP.txtClcb1.Text = "" '材料成本
wbHTP.txtClcb2.Text = ""
wbHTP.txtRgf1.Text = "" '人 工 费
wbHTP.txtRGF2.Text = ""
wbHTP.txtCLF1.Text = "" '差 旅 费
wbHTP.txtCLF2.Text = ""
wbHTP.txtFbje1.Text = "" '分包金额
wbHTP.txtFbje2.Text = ""
wbHTP.txtYf1.Text = "" '运    费
wbHTP.txtYf2.Text = ""
wbHTP.txtYj1.Text = "" '佣    金
wbHTP.txtYj2.Text = ""
wbHTP.txtQt1.Text = "" '项目费用
wbHTP.txtQt2.Text = ""
wbHTP.txtLr1.Text = "" '毛    利
wbHTP.txtLr2.Text = ""
wbHTP.txtJlr1.Text = ""
wbHTP.txtJlr2.Text = ""
wbHTP.optLa.Value = False '增值发票
wbHTP.optLb.Value = False '商业发票
wbHTP.optLc.Value = False '服务发票
wbHTP.txtTc2.Text = "" '提成
wbHTP.txtJy.Text = "" '评审建议
wbHTP.chkA.Caption = ""
wbHTP.lblHid.Caption = ""
wbHTP.chkA.Tag = ""
wbHTP.chkA.Value = 0
wbHTP.chkB.Caption = ""
wbHTP.chkB.Tag = ""
wbHTP.chkB.Value = 0
wbHTP.chkC.Caption = ""
wbHTP.chkC.Value = 0
wbHTP.chkD.Caption = ""
wbHTP.chkD.Value = 0
wbHTP.chkE.Caption = ""
wbHTP.chkE.Value = 0
wbHTP.lblYw.Caption = ""
wbHTP.lblJZ.Caption = ""
wbHTP.lblJl.Caption = ""
wbHTP.lblYZ.Caption = ""
wbHTP.lblZJ.Caption = ""
wbHTP.txtXMNr.Text = ""
wbHTP.txtTcBe.Text = 8 '提成比例
wbHTP.txtTcBe.Visible = False
wbHTP.lblTcBe.Visible = False
wbHTP.UpDa.Visible = False
wbHTP.txtTc2.Visible = True
wbHTP.txtTcRQ.Text = "提成取现日期"


'清空wbMx表
wbMx.txtXdj.Text = "" '巡视单价
wbMx.txtXgT.Text = "" '巡视工时
wbMx.txtXxG.Text = "" '巡视小计
wbMx.txtJdj.Text = "" '急修单价
wbMx.txtJgT.Text = "" '急修工时
wbMx.txtJxG.Text = "" '急修小计
wbMx.txtGdj.Text = "" '工程单价
wbMx.txtGgT.Text = "" '工程工时
wbMx.txtGxG.Text = "" '工程小计
wbMx.txtDdj.Text = "" '大修单价
wbMx.txtDgT.Text = "" '大修工时
wbMx.txtDxG.Text = "" '大修小计
wbMx.txtXgT1.Text = ""
wbMx.txtXxG1.Text = ""
wbMx.txtJgT1.Text = ""
wbMx.txtJxG1.Text = ""
wbMx.txtGgT1.Text = ""
wbMx.txtGxG1.Text = ""
wbMx.txtDgT1.Text = ""
wbMx.txtDxG1.Text = ""

wbMx.LBLhG.Caption = "" '预计人工和计
wbMx.lblHG1.Caption = "" '实际人工和计
wbMx.cmdGzd.Caption = ""
Set wbMx.dtgGzb.DataSource = Nothing


wbMx.txtJPJE.Text = "" '往返机票金额
wbMx.txtJPCou.Text = "" '往返机票人数
wbMx.txtJPXG.Text = "" '往返机票小计
wbMx.txtJPXG1.Text = ""
wbMx.txtHCJE.Text = "" '往返火车票金额
wbMx.txtHCCou.Text = "" '往返火车票人数
wbMx.txtHCXG.Text = "" '往返火车票小计
wbMx.txtHCXG1.Text = ""
wbMx.txtQCJE.Text = "" '往返汽车票金额
wbMx.txtQCCou.Text = "" '往返汽车票人数
wbMx.txtQCXG.Text = "" '往返汽车票小计
wbMx.txtQCXG1.Text = ""
wbMx.txtZJE.Text = "" '住宿金额
wbMx.txtZCou.Text = "" '住宿人数
wbMx.txtZXG.Text = "" '住宿小计
wbMx.txtZXG1.Text = ""
wbMx.txtCJE.Text = "" '餐费金额
wbMx.txtCCou.Text = "" '餐费人数
wbMx.txtCXG.Text = "" '餐费小计
wbMx.txtCXG1.Text = ""
wbMx.txtDDJE.Text = "" '当地车费
wbMx.txtDDCou.Text = "" '当地车费人数
wbMx.txtDDXG.Text = "" '当地车费小计
wbMx.txtDDXG1.Text = ""


wbMx.lblCf.Caption = "" '预计差旅费和计
wbMx.lblCF1.Caption = "" '实际差旅费和计
Set wbMx.dtgCl.DataSource = Nothing

End Sub





Public Sub wbAdd() '维保合同评审单数据添加


End Sub




































































Public Sub wbBound() '维保合同评审单字段绑定
On Error Resume Next
Dim tt As String
Dim xZ As String

'记录打开日志
Call mod1.zhuDa(1, mod1.HTP.Fields("htbh").Value)
wbHTP.txtKhmc.Text = mod1.HTP.Fields("khmc").Value '客户名称
wbHTP.txtXmmc.Text = mod1.HTP.Fields("xmmc").Value '项目名称
wbHTP.txtYwy.Text = mod1.HTP.Fields("Ywy").Value '业务员
wbHTP.txtHtbh.Text = mod1.HTP.Fields("htBh").Value '合同编号
wbHTP.txtGLG.Text = mod1.HTP.Fields("GLG").Value '管理公司
wbHTP.txtADR.Text = mod1.HTP.Fields("khADR").Value '项目地址
wbHTP.txtHtdate.Text = Format(mod1.HTP.Fields("htRq").Value, "Long Date") '合同日期
wbHTP.dt1.Value = mod1.HTP.Fields("htRq").Value
xZ = mod1.HTP.Fields("htXz").Value '合同性质
wbHTP.lblHid.Caption = mod1.HTP.Fields("hid").Value
Select Case xZ
'Case "A. 零配件合同"
'wbHTP.optA(0).Value = True
'Case "B1.工程合同"
'wbHTP.optA(1).Value = True
Case "C. 维保合同"
wbHTP.optA(3).Value = True
Case "D. 维修合同"
wbHTP.optA(4).Value = True
'Case "E. 产品合同"
'wbHTP.optA(5).Value = True
End Select
wbHTP.txtKhdm.Text = mod1.HTP.Fields("khDh").Value '客户代号
wbHTP.comQy.Text = mod1.HTP.Fields("qy").Value '区域

'合同期限起始期
wbHTP.dt3.Value = mod1.HTP.Fields("htQy").Value
'合同期限结束期
wbHTP.dt4.Value = mod1.HTP.Fields("htQy1").Value

wbHTP.txtMOn.Text = mod1.HTP.Fields("bxQ").Value '保修期

If mod1.HTP.Fields("htZe").Value = 0 Then
wbHTP.txtHtze.Text = ""
Else
wbHTP.txtHtze.Text = mod1.HTP.Fields("htZe").Value '合同总额
wbMx.lblHtze.Caption = wbHTP.txtHtze.Text
End If

wbMx.txtFkBz.Text = mod1.HTP.Fields("fkBz").Value '付款条件备注

If mod1.HTP.Fields("cbZe").Value = 0 Then
wbHTP.txtCbze1.Text = ""
Else
wbHTP.txtCbze1.Text = mod1.HTP.Fields("cbZe").Value '成本总额
End If

If mod1.HTP.Fields("cbze1").Value = 0 Then
wbHTP.txtCbze2.Text = ""
Else
wbHTP.txtCbze2.Text = mod1.HTP.Fields("cbze1").Value '实际成本总额
End If


wbHTP.txtClcb1.Text = mod1.HTP.Fields("clCb").Value '材料成本



wbHTP.txtClcb2.Text = mod1.HTP.Fields("clCb1").Value '实际材料成本




wbHTP.txtRgf1.Text = mod1.HTP.Fields("rgF").Value '人工费


wbHTP.txtRGF2.Text = mod1.HTP.Fields("rgF1").Value '实际人工费


'If mod1.HtP.Fields("clF1").Value = 0 Then '差旅费
'wbHTP.txtCLF1.Text = ""
'Else
wbHTP.txtCLF1.Text = mod1.HTP.Fields("clF1").Value
'End If
If mod1.HTP.Fields("clF2").Value = 0 Then '实际差旅费
wbHTP.txtCLF2.Text = ""
Else
wbHTP.txtCLF2.Text = mod1.HTP.Fields("clF2").Value
End If


If mod1.HTP.Fields("fbJe").Value = 0 Then
wbHTP.txtFbje1.Text = ""
Else
wbHTP.txtFbje1.Text = mod1.HTP.Fields("fbJe").Value '分包金额
End If

If mod1.HTP.Fields("fbJe1").Value = 0 Then
wbHTP.txtFbje2.Text = ""
Else
wbHTP.txtFbje2.Text = mod1.HTP.Fields("fbJe1").Value '实际分包金额
End If


If mod1.HTP.Fields("yunF").Value = 0 Then
wbHTP.txtYf1.Text = ""
Else
wbHTP.txtYf1.Text = mod1.HTP.Fields("yunF").Value '运费
End If

If mod1.HTP.Fields("yunF1").Value = 0 Then
wbHTP.txtYf2.Text = ""
Else
wbHTP.txtYf2.Text = mod1.HTP.Fields("yunF1").Value '实际运费
End If

'If mod1.HtP.Fields("Yj").Value = 0 Then
'wbHTP.txtYj1.Text = ""
'Else
wbHTP.txtYj1.Text = mod1.HTP.Fields("Yj").Value '佣金
'End If

If mod1.HTP.Fields("Yj1").Value = 0 Then
wbHTP.txtYj2.Text = ""
Else
wbHTP.txtYj2.Text = mod1.HTP.Fields("Yj1").Value '实际佣金
End If


'If mod1.HtP.Fields("qtF").Value = 0 Then
'wbHTP.txtQt1.Text = ""
'Else
wbHTP.txtQt1.Text = mod1.HTP.Fields("qtF").Value '项目费用
'End If

If mod1.HTP.Fields("qtF1").Value = 0 Then
wbHTP.txtQt2.Text = ""
Else
wbHTP.txtQt2.Text = mod1.HTP.Fields("qtF1").Value '实际项目费用
End If


If mod1.HTP.Fields("xmLr").Value = 0 Then
wbHTP.txtLr1.Text = ""
Else
wbHTP.txtLr1.Text = mod1.HTP.Fields("xmLr").Value '项目利润
End If

If mod1.HTP.Fields("xmLr1").Value = 0 Then
wbHTP.txtLr2.Text = ""
Else
wbHTP.txtLr2.Text = mod1.HTP.Fields("xmLr1").Value '实际项目利润
End If

wbHTP.txtJlr1.Text = mod1.HTP.Fields("jlr1").Value
wbHTP.txtJlr2.Text = mod1.HTP.Fields("jlr2").Value

'发票类型
If mod1.HTP.Fields("fpLx").Value = "增值发票" Then
wbHTP.optLa.Value = True
ElseIf mod1.HTP.Fields("fpLx").Value = "商业发票" Then
wbHTP.optLb.Value = True
ElseIf mod1.HTP.Fields("fpLx").Value = "服务发票" Then
wbHTP.optLc.Value = True
End If


If mod1.HTP.Fields("Tc1").Value = 0 Then
wbHTP.txtTc2.Text = ""
Else
wbHTP.txtTc2.Text = mod1.HTP.Fields("Tc1").Value '实际提成
End If

wbHTP.txtTcBe.Text = mod1.HTP.Fields("TCbe").Value '提成比例


wbHTP.txtJy.Text = mod1.HTP.Fields("jy").Value '评审建议

If mod1.HTP.Fields("ywQz").Value <> "" Then '业务员签字
wbHTP.chkA.Caption = mod1.HTP.Fields("ywQz").Value
wbHTP.chkA.Tag = mod1.HTP.Fields("xywy").Value
'wbHTP.chkA.Value = 1
Else
wbHTP.chkA.Caption = ""
wbHTP.chkA.Value = 0
End If
If mod1.HTP.Fields("JzQz").Value <> "" Then '技术支持签字
wbHTP.chkE.Caption = mod1.HTP.Fields("JzQz").Value
'wbHTP.chkE.Value = 1
Else
wbHTP.chkE.Caption = ""
wbHTP.chkE.Value = 0
End If
If mod1.HTP.Fields("jlQz").Value <> "" Then '销售经理签字
wbHTP.chkB.Caption = mod1.HTP.Fields("jlQz").Value
wbHTP.chkB.Tag = mod1.HTP.Fields("xjlq").Value
'wbHTP.chkB.Value = 1
Else
wbHTP.chkB.Caption = ""
wbHTP.chkB.Value = 0
End If
If mod1.HTP.Fields("yzQz").Value <> "" Then '商务经理签字
wbHTP.chkC.Caption = mod1.HTP.Fields("yzQz").Value
'wbHTP.chkC.Value = 1
Else
wbHTP.chkC.Caption = ""
wbHTP.chkC.Value = 0
End If
If mod1.HTP.Fields("zjQz").Value <> "" Then '总经理签字
wbHTP.chkD.Caption = mod1.HTP.Fields("zjQz").Value
'wbHTP.chkD.Value = 1
Else
wbHTP.chkD.Caption = ""
wbHTP.chkD.Value = 0
End If

wbHTP.lblYw.Caption = mod1.HTP.Fields("ywDa").Value '销售签字日期
wbHTP.lblJZ.Caption = mod1.HTP.Fields("JzDa").Value '技术支持签字日期
wbHTP.lblJl.Caption = mod1.HTP.Fields("JlDa").Value '销售经理签字日期
wbHTP.lblYZ.Caption = mod1.HTP.Fields("YzDa").Value '商务经理签字日期
wbHTP.lblZJ.Caption = mod1.HTP.Fields("ZjDa").Value '总经理签字日期
wbHTP.txtXMNr.Text = mod1.HTP.Fields("xmnr").Value '

If mod1.HTP.Fields("jTf").Value = True Then
    wbHTP.cmdCount.Caption = "已结算"
    wbHTP.cmdCount.Enabled = False
Else
    wbHTP.cmdCount.Caption = "计算"
    wbHTP.cmdCount.Enabled = True
  
End If


If IsNull(mod1.HTP.Fields("TCRQ").Value) = False Then
    wbHTP.txtTcRQ.Text = mod1.HTP.Fields("TCRQ").Value '提成取现日期
End If

'如果为旧合同,则评审阶段设为蓝色
If mod1.HTP.Fields("XGG").Value = True Then
    wbHTP.optP.BackColor = &HC0FFFF
ElseIf mod1.HTP.Fields("XGG").Value = 0 Then
    wbHTP.optP.BackColor = &H8000000F
End If

wbHTP.optP.Enabled = False
wbHTP.optG.Enabled = False
wbHTP.optZ.Enabled = False
wbHTP.optW.Enabled = False

'合同执行否
If mod1.HTP.Fields(51).Value = 0 Then
    wbHTP.optP.Value = True
'    wbHTP.optP.Enabled = True
'    '如果都签字了,而且为盖章者打开,则可以进行盖章
'    If (wbHTP.chkC.Caption <> "" And wbHTP.txtHtze.Text < 10000 Or wbHTP.chkD.Caption <> "") And _
'        mod1.KGZ = True Then
'        wbHTP.optG.Enabled = True
'    End If
    
ElseIf mod1.HTP.Fields("htF").Value = 9 Then
    wbHTP.optG.Value = True
'    wbHTP.optG.Enabled = True
'    If mod1.KZX = True Then
'        wbHTP.optZ.Enabled = True
'    End If

ElseIf mod1.HTP.Fields(51).Value = 1 Then
    wbHTP.optZ.Value = True
'    wbHTP.optZ.Enabled = True
'    If mod1.KWC = True Then
'        wbHTP.optW.Enabled = True
'    End If

ElseIf mod1.HTP.Fields(51).Value = 2 Then
    wbHTP.optW.Value = True
    'wbHTP.optW.Enabled = True
End If



wbHTP.lblBM.Caption = mod1.HTP.Fields("bm").Value

''如果合同评审没有全通过或者不是小吴，则“合同执行否”字段不能编辑
'If wbHTP.chkD.Caption <> "" And frmLogin.Combo1.Text = "胡颖" Then
'wbHTP.frmZt.Enabled = True
'Else

'End If

'绑定wbRGMX表
Set adoRGF = CreateObject("adodb.recordset")
tt = "Select * from wbRGMX where htBh='" & wbHTP.txtHtbh.Text & "'"
adoRGF.Close
adoRGF.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
wbMx.txtXdj.Text = adoRGF.Fields("Xdj").Value '巡视单价
wbMx.txtXgT.Text = adoRGF.Fields("XgT").Value '巡视工时
wbMx.txtXxG.Text = adoRGF.Fields("XxG").Value '巡视小计
wbMx.txtJdj.Text = adoRGF.Fields("Jdj").Value '急修单价
wbMx.txtJgT.Text = adoRGF.Fields("JgT").Value '急修工时
wbMx.txtJxG.Text = adoRGF.Fields("JxG").Value '急修小计
wbMx.txtGdj.Text = adoRGF.Fields("Gdj").Value '工程单价
wbMx.txtGgT.Text = adoRGF.Fields("GgT").Value '工程工时
wbMx.txtGxG.Text = adoRGF.Fields("GxG").Value '工程小计
wbMx.txtDdj.Text = adoRGF.Fields("Ddj").Value '大修单价
wbMx.txtDgT.Text = adoRGF.Fields("DgT").Value '大修工时
wbMx.txtDxG.Text = adoRGF.Fields("DxG").Value '大修小计
'预计人工和计
wbMx.LBLhG.Caption = wbHTP.txtRgf1.Text
wbMx.txtJPJE.Text = adoRGF.Fields("JPJE").Value '往返机票金额
wbMx.txtJPCou.Text = adoRGF.Fields("JPCou").Value '往返机票人数
wbMx.txtJPXG.Text = adoRGF.Fields("JPXG").Value '往返机票小计
wbMx.txtHCJE.Text = adoRGF.Fields("HCJE").Value '往返火车票金额
wbMx.txtHCCou.Text = adoRGF.Fields("HCCou").Value '往返火车票人数
wbMx.txtHCXG.Text = adoRGF.Fields("HCXG").Value '往返火车票小计
wbMx.txtQCJE.Text = adoRGF.Fields("QCJE").Value '往返汽车票金额
wbMx.txtQCCou.Text = adoRGF.Fields("QCCou").Value '往返汽车票人数
wbMx.txtQCXG.Text = adoRGF.Fields("QCXG").Value '往返汽车票小计
wbMx.txtZJE.Text = adoRGF.Fields("ZJE").Value '住宿金额
wbMx.txtZCou.Text = adoRGF.Fields("ZCou").Value '住宿人数
wbMx.txtZXG.Text = adoRGF.Fields("ZXG").Value '住宿小计
wbMx.txtCJE.Text = adoRGF.Fields("CJE").Value '餐费金额
wbMx.txtCCou.Text = adoRGF.Fields("CCou").Value '餐费人数
wbMx.txtCXG.Text = adoRGF.Fields("CXG").Value '餐费小计
wbMx.txtDDJE.Text = adoRGF.Fields("DDJE").Value '当地车费
wbMx.txtDDCou.Text = adoRGF.Fields("DDCou").Value '当地车费人数
wbMx.txtDDXG.Text = adoRGF.Fields("DDXG").Value '当地车费小计
'预计差旅费和计
wbMx.lblCf.Caption = wbHTP.txtCLF1.Text

'绑定佣金表
tt = "Select * from Yongjin where htBh='" & wbHTP.txtHtbh.Text & "' order by yId"
frmYJ.adoYj.Recordset.Close
frmYJ.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Set frmYJ.dtgYJ.DataSource = frmYJ.adoYj

wbHTP.frmZt.Enabled = True
End Sub






Public Sub GXZJ(Htbh As String)  '计算购销合同的实际成本总额及实际利润和提成
Dim tt As String
On Error Resume Next
'计算实际成本总额（材料成本+预留费用+运费+佣金）
tt = "update htping set cbZe1=clcb1+qtF1+yunF1+Yj1 where htbh='" & Htbh & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'计算实际利润
tt = "update htping set xmLr1=htZe-cbZe1 where htbh='" & Htbh & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'计算实际提成
tt = "update htping set Tc1=xmLr1*0.08 where htbh='" & Htbh & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
End Sub

Public Sub gxMr() '默认的购销合同条款
htgX.txtT2.Text = "按厂家生产标准"
'htgX.txtT4.Text = "供方发货"
htgX.txtT5.Text = "需方负担"
htgX.txtT7.Text = "标准包装"
htgX.txtT8.Text = "货到现场验明数量，如有异议三天内提出"
'htgX.txtT10.Text = "款到发货"
htgX.txtT12.Text = "按照《中华人民共和国合同法》"
htgX.txtT13.Text = "协商、仲裁、诉讼（上海）"
End Sub

Public Sub NewBound(Hid As Long)
Dim tt As String
On Error Resume Next
frmWbNew.frmNb.Visible = False
frmWbNew.frmDx.Visible = False

tt = "select * from htping where hid=" & Hid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
frmWbNew.Visible = False
frmWbNew.txtKhmc.Text = mod1.HTP.Fields("khmc").Value
frmWbNew.txtKhdm.Text = mod1.HTP.Fields("khdh").Value
frmWbNew.txtHtbh.Text = mod1.HTP.Fields("htbh").Value
frmWbNew.lblHtxz.Caption = mod1.HTP.Fields("htxz").Value
frmWbNew.tabGc.TabCaption(2) = mod1.HTP.Fields("htxz").Value
frmWbNew.txtXmmc.Text = mod1.HTP.Fields("xmmc").Value
frmWbNew.txtXmmc.Tag = mod1.HTP.Fields("xid").Value
frmWbNew.txtXYwy.Text = mod1.HTP.Fields("Xywy").Value
frmWbNew.txtXYwy.Tag = mod1.HTP.Fields("xuid").Value
frmWbNew.txtHtrq.Text = mod1.HTP.Fields("htrq").Value
frmWbNew.comQy.Text = mod1.HTP.Fields("qy").Value
frmWbNew.txtADR.Text = mod1.HTP.Fields("khadr").Value
frmWbNew.txtHtze.Text = mod1.HTP.Fields("htze").Value
'frmWbNew.txtZe.Text = mod1.HTP.Fields("htbh").Value  '财务收款
'frmWbNew.txtEd.Text = mod1.HTP.Fields("htbh").Value
'开票类型

If mod1.HTP.Fields("fpLX").Value = "增值发票" Then
    frmWbNew.optLa.Value = True
ElseIf mod1.HTP.Fields("fpLX").Value = "商业发票" Then
    frmWbNew.optLb.Value = True
ElseIf mod1.HTP.Fields("fpLX").Value = "服务发票" Then
    frmWbNew.optLc.Value = True
End If
frmWbNew.txtBz.Text = mod1.HTP.Fields("bz").Value
frmWbNew.txtFbje1.ToolTipText = mod1.HTP.Fields("fbnr").Value
frmWbNew.txtFbje2.ToolTipText = mod1.HTP.Fields("fbnr").Value
frmWbNew.txtCbze1.Text = mod1.HTP.Fields("cbze").Value
frmWbNew.txtClcb1.Text = mod1.HTP.Fields("clCb").Value
frmWbNew.txtRgf1.Text = mod1.HTP.Fields("rgF").Value
frmWbNew.txtCLF1.Text = mod1.HTP.Fields("clf1").Value
frmWbNew.txtYf1.Text = mod1.HTP.Fields("yunF").Value
frmWbNew.txtQt1.Text = mod1.HTP.Fields("qtF1").Value
frmWbNew.txtJlr1.Text = mod1.HTP.Fields("Jlr1").Value
frmWbNew.txtYj1.Text = mod1.HTP.Fields("Yj").Value
frmWbNew.txtLr1.Text = mod1.HTP.Fields("xmLr").Value
frmWbNew.txtTcBe.Text = mod1.HTP.Fields("tcBe").Value
frmWbNew.txtTc2.Text = mod1.HTP.Fields("Tc1").Value
frmWbNew.txtTcRQ.Text = mod1.HTP.Fields("TCRQ").Value

frmWbNew.txtCbze2.Text = mod1.HTP.Fields("cbze1").Value
frmWbNew.txtFbje1.Text = mod1.HTP.Fields("fbje").Value
frmWbNew.txtFbje2 = mod1.HTP.Fields("fbje1").Value
frmWbNew.txtYf2 = mod1.HTP.Fields("yunF1").Value
frmWbNew.txtQt2 = mod1.HTP.Fields("qtF").Value '已经发生的项目费用


frmWbNew.txtJlr2 = mod1.HTP.Fields("Jlr2").Value
frmWbNew.txtYj2 = mod1.HTP.Fields("Yj1").Value
frmWbNew.txtLr2 = mod1.HTP.Fields("xmLr1").Value
frmWbNew.lblyjFF.Caption = mod1.HTP.Fields("yjff").Value

If mod1.HTP.Fields("htqy").Value = "1999-1-1" Or IsNull(mod1.HTP.Fields("htqy").Value) = True Then
    frmWbNew.txtF.Text = ""
Else
    frmWbNew.txtF.Text = mod1.HTP.Fields("htqy").Value
End If
If mod1.HTP.Fields("htqy1").Value = "1999-1-1" Or IsNull(mod1.HTP.Fields("htqy1").Value) = True Then
    frmWbNew.txtL.Text = ""
Else
    frmWbNew.txtL.Text = mod1.HTP.Fields("htqy1").Value
End If
If frmWbNew.txtL.Text = frmWbNew.txtF.Text Then
    frmWbNew.txtL.Text = ""
    frmWbNew.txtF.Text = ""
End If
Set frmWbNew.dtgFk.DataSource = Nothing
Set frmWbNew.dtgYf.DataSource = Nothing

If mod1.HTP.Fields("htf").Value = 0 Then
    frmWbNew.optP.Value = True
ElseIf mod1.HTP.Fields("htf").Value = 1 Then
    frmWbNew.optZ.Value = True
ElseIf mod1.HTP.Fields("htf").Value = 9 Then
    frmWbNew.optG.Value = True
ElseIf mod1.HTP.Fields("htf").Value = 2 Then
    frmWbNew.optW.Value = True
End If

If mod1.HTP.Fields("htf").Value = 1 And mod1.DName = mod1.HTP.Fields("xywy").Value Then
    frmWbNew.cmdKP.Visible = True
End If

frmWbNew.lblHid.Caption = mod1.HTP.Fields("hid").Value
frmWbNew.lblBaoId.Caption = mod1.HTP.Fields("baoid").Value
frmWbNew.lblPwf.Caption = mod1.HTP.Fields("pwf").Value
frmWbNew.lblLc.Caption = mod1.HTP.Fields("Lc").Value
frmWbNew.lblLcRen.Caption = mod1.HTP.Fields("LcRen").Value
frmWbNew.lblLcUid.Caption = mod1.HTP.Fields("LcUid").Value
frmWbNew.lblFwid.Caption = mod1.HTP.Fields("Fwid").Value
frmWbNew.lblNlb.Caption = mod1.HTP.Fields("Nlb").Value
frmWbNew.lblLcou.Caption = mod1.HTP.Fields("Lcou").Value


frmWbNew.lblYwy.Caption = mod1.HTP.Fields("xywy").Value
frmWbNew.lblUid.Caption = mod1.HTP.Fields("xuid").Value

'财务评定
'业绩
'frmWbNew.txtYjf.Text = mod1.HTP.Fields("yjRQ").Value
If mod1.HTP.Fields("yjrq").Value = "1999-1-1" Then frmWbNew.txtYjf.Text = ""
If mod1.HTP.Fields("yjf").Value = True Then
    frmWbNew.chkYJF.Value = 1
Else
    frmWbNew.chkYJF.Value = 0
End If
If mod1.HTP.Fields("jtf").Value = True Then
    frmWbNew.chkJTF.Value = 1
Else
    frmWbNew.chkJTF.Value = 0
End If
If mod1.HTP.Fields("qkf").Value = True Then
    frmWbNew.chkQKF.Value = 1
Else
    frmWbNew.chkQKF.Value = 0
End If
frmWbNew.txtYjfBz.Text = mod1.HTP.Fields("yjbz").Value
Set mod1.mJt = CreateObject("adodb.recordset")
Set mod1.mQk = CreateObject("adodb.recordset")
Set mod1.mYjF = CreateObject("adodb.recordset")

tt = "select rq as 日期,je as 金额,bz as 备注,mid from htpingJt where hid=" & Val(frmWbNew.lblHid.Caption) & " and delf=1 order by mid desc"
mod1.mJt.Close
mod1.mJt.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.mJt.RecordCount) = True Then
    MsgBox ("读取数据错误2.2!")
    Exit Sub
End If

If mod1.mJt.RecordCount = 0 Then
    Set frmWbNew.dtgJTf.DataSource = mod1.mJt
    frmWbNew.dtgJTf.Rows = 2
    frmWbNew.dtgJTf.FixedRows = 0
    frmWbNew.dtgJTf.FixedRows = 1
Else
    frmWbNew.dtgJTf.Rows = 2
    frmWbNew.dtgJTf.FixedRows = 1
    Set frmWbNew.dtgJTf.DataSource = mod1.mJt
End If

tt = "select sum(je) as je from htpingJt where hid=" & Val(frmWbNew.lblHid.Caption) & " and delf=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.HTP.RecordCount) = True Then
    MsgBox ("读取数据错误2.3!")
    Exit Sub
End If
frmWbNew.txtJTf.Text = mod1.HTP.Fields("je").Value

tt = "select rq as 日期,je as 金额,bz as 备注,mid from htpingQk where hid=" & Val(frmWbNew.lblHid.Caption) & " and delf=1 order by mid desc"
mod1.mYjF.Close
mod1.mYjF.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.mYjF.RecordCount) = True Then
    MsgBox ("读取数据错误2.8!")
    Exit Sub
End If
If mod1.mYjF.RecordCount = 0 Then
    Set frmWbNew.dtgyjF.DataSource = mod1.mYjF
    frmWbNew.dtgyjF.Rows = 2
    frmWbNew.dtgyjF.FixedRows = 0
    frmWbNew.dtgyjF.FixedRows = 1
Else
    frmWbNew.dtgyjF.Rows = 2
    frmWbNew.dtgyjF.FixedRows = 1
    Set frmWbNew.dtgyjF.DataSource = mod1.mYjF
End If
tt = "select sum(je) as je from htpingQk where hid=" & Val(frmWbNew.lblHid.Caption) & " and delf=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.HTP.RecordCount) = True Then
    MsgBox ("读取数据错误2.6!")
    Exit Sub
End If
frmWbNew.txtYjf.Text = mod1.HTP.Fields("je").Value



tt = "SELECT rp_dd as 日期,amtn_cls as 金额,rem as 备注 FROM TF_MON where rp_id=1 and cas_no='" & frmWbNew.txtHtbh.Text & "' order by rp_dd"
mod1.mQk.Close
mod1.mQk.Open tt, mod1.workTx, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.mQk.RecordCount) = True Then
    MsgBox ("读取数据错误2.5!")
    Exit Sub
End If

If mod1.mQk.RecordCount = 0 Then
    Set frmWbNew.dtgQkf.DataSource = mod1.mQk
    frmWbNew.dtgQkf.Rows = 2
    frmWbNew.dtgQkf.FixedRows = 0
    frmWbNew.dtgQkf.FixedRows = 1
Else
    frmWbNew.dtgQkf.Rows = 2
    frmWbNew.dtgQkf.FixedRows = 1
    Set frmWbNew.dtgQkf.DataSource = mod1.mQk
End If

'tt = "select sum(je) as je from htpingQk where hid=" & Val(frmWbNew.lblHid.Caption) & " and delf=1"
tt = "select sum(amtn_cls) as je from tf_mon where rp_id=1 and cas_no='" & frmWbNew.txtHtbh.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workTx, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.HTP.RecordCount) = True Then
    MsgBox ("读取数据错误2.6!")
    Exit Sub
End If
frmWbNew.txtQkf.Text = mod1.HTP.Fields("je").Value
frmWbNew.txtZe.Text = frmWbNew.txtQkf.Text
frmWbNew.txtEd.Text = Round(Val(frmWbNew.txtZe.Text) / Val(frmWbNew.txtHtze.Text) * 100, 2)




tt = "select qy,bm from renyuan where username='" & frmWbNew.lblYwy.Caption & "' and userid='" & frmWbNew.lblUid.Caption & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
frmWbNew.lblBM.Caption = mod1.HTP.Fields("bm").Value
frmWbNew.lblQy.Caption = mod1.HTP.Fields("qy").Value
frmWbNew.comQy.Text = mod1.HTP.Fields("qy").Value

frmWbNew.frmJi.Visible = True



If frmWbNew.lblHtxz.Caption = "维保" Then
    tt = "select zh,zName,jzpb,jzxh,sl,ta,tb,tc,mon,wc,xc,dxnr,bid from baoJiaD where baoid=" & Val(frmWbNew.lblBaoId.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    frmWbNew.Bid = mod1.HTP.Fields("bid").Value
    frmWbNew.comZu.Text = mod1.HTP.Fields("zh").Value
    frmWbNew.txtZu.Text = mod1.HTP.Fields("zName").Value
    frmWbNew.comPb.Text = mod1.HTP.Fields("jzpb").Value
    frmWbNew.comXh.Text = mod1.HTP.Fields("jzxh").Value
    frmWbNew.txtSl.Text = mod1.HTP.Fields("sl").Value
    If mod1.HTP.Fields("ta").Value = True Then
        frmWbNew.chkBa.Value = 1
    Else
        frmWbNew.chkBa.Value = 0
    End If
    If mod1.HTP.Fields("tb").Value = True Then
        frmWbNew.chkBb.Value = 1
    Else
        frmWbNew.chkBb.Value = 0
    End If
    If mod1.HTP.Fields("tc").Value = True Then
        frmWbNew.chkBc.Value = 1
    Else
        frmWbNew.chkBc.Value = 0
    End If
    frmWbNew.txtMOn.Text = mod1.HTP.Fields("mon").Value
    frmWbNew.txtWc.Text = mod1.HTP.Fields("wc").Value
    frmWbNew.txtXc.Text = mod1.HTP.Fields("xc").Value
    frmWbNew.cmdTk.Visible = False
    If frmWbNew.comPb.Text = "" Then
        tt = "select jzpb as 机组品牌,jzxh as 机组型号,sl as 数量,jxId from wbjb where baoid=" & Val(frmWbNew.lblBaoId.Caption)
        Set frmWbNew.adoA = CreateObject("adodb.recordset")
        frmWbNew.adoA.Close
        frmWbNew.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmWbNew.dtgA.DataSource = frmWbNew.adoA
        frmWbNew.dtgA.Visible = True
        frmWbNew.cmdTk.Visible = True
    Else
        frmWbNew.cmdTk.Visible = False
        frmWbNew.dtgA.Visible = False
        '年保
        tt = "select * from xunJIaWbView where wbx='年保' and bid=" & mod1.HTP.Fields("bid").Value
        frmWbNew.adoWb.Close
        frmWbNew.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmWbNew.dtgWb.DataSource = frmWbNew.adoWb
        frmWbNew.dtgWb.FixedRows = 0
        frmWbNew.dtgWb.MergeCol(1) = True
        frmWbNew.dtgWb.MergeCol(2) = True
        frmWbNew.dtgWb.MergeCol(3) = True
        frmWbNew.dtgWb.MergeCells = 3
        frmWbNew.dtgWb.FixedRows = 1
        '例检表
        tt = "select * from xunJIaWbView where wbx='例检' and bid=" & mod1.HTP.Fields("bid").Value
        frmWbNew.adoLj.Close
        frmWbNew.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmWbNew.dtgLj.DataSource = frmWbNew.adoLj
        frmWbNew.dtgLj.FixedRows = 0
        frmWbNew.dtgLj.MergeCol(1) = True
        frmWbNew.dtgLj.MergeCol(2) = True
        frmWbNew.dtgLj.MergeCol(3) = True
        frmWbNew.dtgLj.MergeCells = 3
        frmWbNew.dtgLj.FixedRows = 1
    End If
    

    
    '显示产品列表
    tt = "select * from BaoJiaMxView where baoid=" & Val(frmWbNew.lblBaoId.Caption) & " order by lid"
    frmWbNew.adoBx.Close
    frmWbNew.adoBx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbNew.dtgBao.DataSource = frmWbNew.adoBx
    frmWbNew.dtgBao.FixedRows = 0
    frmWbNew.dtgBao.MergeCol(1) = True
    frmWbNew.dtgBao.MergeCol(2) = True
    frmWbNew.dtgBao.MergeCol(10) = True
    frmWbNew.dtgBao.MergeCol(14) = True
    frmWbNew.dtgBao.MergeCells = 3
    frmWbNew.dtgBao.FixedRows = 1
    '显示成本表
    tt = "select * from xunJiaMxView where bid=0"
    frmWbNew.adoGx.Close
    frmWbNew.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbNew.dtgMa.DataSource = frmWbNew.adoGx
    frmWbNew.frmTime.Visible = True
    frmWbNew.frmNb.Visible = True
    frmWbNew.tabGc.TabVisible(0) = True
    frmWbNew.tabGc.TabVisible(1) = True
    frmWbNew.tabGc.TabVisible(2) = False
    frmWbNew.tabGc.TabVisible(3) = True
    frmWbNew.tabGc.Tab = 0
    frmWbNew.dtgWb.Visible = True
ElseIf frmWbNew.lblHtxz.Caption = "大修" Or frmWbNew.lblHtxz.Caption = "工程分包" Then
    tt = "select zh,zName,jzpb,jzxh,sl,ta,tb,tc,mon,wc,xc,dxnr,bid from baoJiaD where baoid=" & Val(frmWbNew.lblBaoId.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    frmWbNew.comZu.Text = mod1.HTP.Fields("zh").Value
    frmWbNew.txtZu.Text = mod1.HTP.Fields("zName").Value
    frmWbNew.comPb.Text = mod1.HTP.Fields("jzpb").Value
    frmWbNew.comXh.Text = mod1.HTP.Fields("jzxh").Value
    frmWbNew.txtSl.Text = mod1.HTP.Fields("sl").Value
    frmWbNew.chkBa.Value = mod1.HTP.Fields("ta").Value
    frmWbNew.chkBb.Value = mod1.HTP.Fields("tb").Value
    frmWbNew.chkBc.Value = mod1.HTP.Fields("tc").Value
    frmWbNew.txtMOn.Text = mod1.HTP.Fields("mon").Value
    frmWbNew.txtWc.Text = mod1.HTP.Fields("wc").Value
    frmWbNew.txtXc.Text = mod1.HTP.Fields("xc").Value
    frmWbNew.txtDXNR.Text = mod1.HTP.Fields("dxnr").Value
    frmWbNew.cmdTk.Visible = False
    If frmWbNew.comPb.Text = "" Then
        tt = "select jzpb as 机组品牌,jzxh as 机组型号,sl as 数量,jxId from wbjb where baoid=" & Val(frmWbNew.lblBaoId.Caption)
        Set frmWbNew.adoA = CreateObject("adodb.recordset")
        frmWbNew.adoA.Close
        frmWbNew.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmWbNew.dtgA.DataSource = frmWbNew.adoA
        frmWbNew.dtgA.Visible = True
    Else
        frmWbNew.dtgA.Visible = False
    End If
    
    '显示产品列表
    tt = "select * from BaoJiaMxView where baoid=" & Val(frmWbNew.lblBaoId.Caption) & " order by lid"
    frmWbNew.adoBx.Close
    frmWbNew.adoBx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbNew.dtgBao.DataSource = frmWbNew.adoBx
    frmWbNew.dtgBao.FixedRows = 0
    frmWbNew.dtgBao.MergeCol(1) = True
    frmWbNew.dtgBao.MergeCol(2) = True
    frmWbNew.dtgBao.MergeCol(10) = True
    frmWbNew.dtgBao.MergeCol(14) = True
    frmWbNew.dtgBao.MergeCells = 3
    frmWbNew.dtgBao.FixedRows = 1
    '显示成本表
    tt = "select * from xunJiaMxView where bid=0"
    frmWbNew.adoGx.Close
    frmWbNew.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbNew.dtgMa.DataSource = frmWbNew.adoGx
    frmWbNew.frmDx.Visible = True
    frmWbNew.tabGc.TabVisible(0) = False
    frmWbNew.tabGc.TabVisible(1) = False
    frmWbNew.tabGc.TabVisible(2) = True
    frmWbNew.tabGc.TabVisible(3) = True
    frmWbNew.frmTime.Visible = False
    frmWbNew.tabGc.Tab = 2
    frmWbNew.txtDXNR.Visible = True
ElseIf frmWbNew.lblHtxz.Caption = "产品" Or frmWbNew.lblHtxz.Caption = "零配件" Or frmWbNew.lblHtxz.Caption = "工程分包" Or frmWbNew.lblHtxz.Caption = "购销" Then
    '显示产品列表
    tt = "select * from BaoJiaMxView where baoid=" & Val(frmWbNew.lblBaoId.Caption) & " order by lid"
    frmWbNew.adoBx.Close
    frmWbNew.adoBx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbNew.dtgBao.DataSource = frmWbNew.adoBx
    frmWbNew.dtgBao.FixedRows = 0
    frmWbNew.dtgBao.MergeCol(1) = True
    frmWbNew.dtgBao.MergeCol(2) = True
    frmWbNew.dtgBao.MergeCol(10) = True
    frmWbNew.dtgBao.MergeCol(14) = True
    frmWbNew.dtgBao.MergeCells = 3
    frmWbNew.dtgBao.FixedRows = 1
    '显示成本表
    tt = "select * from xunJiaMxView where bid=0"
    frmWbNew.adoGx.Close
    frmWbNew.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbNew.dtgMa.DataSource = frmWbNew.adoGx
    frmWbNew.tabGc.TabVisible(0) = False
    frmWbNew.tabGc.TabVisible(1) = False
    frmWbNew.tabGc.TabVisible(2) = False
    frmWbNew.tabGc.TabVisible(3) = True
    frmWbNew.frmJi.Visible = False
    frmWbNew.dtgBao.Visible = True
    frmWbNew.dtgMa.Visible = True
End If

'显示固定费用
tt = "select lb as 费用类别,year(nd) as 年度,qdj as 单价,rl as 人数,xg as 小计,baoid,hid,gid from xmgd where hid=" & Val(frmWbNew.lblHid.Caption)
frmWbNew.adoGD.Close
frmWbNew.adoGD.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWbNew.dtgGD.DataSource = frmWbNew.adoGD
tt = "select sum(xg) as xg from xmgd where baoid=" & Val(frmWbNew.lblBaoId.Caption)
frmWbNew.adoHGD.Close
frmWbNew.adoHGD.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
frmWbNew.txtGd.Text = frmWbNew.adoHGD.Fields("xg").Value
'frmWbNew.txtXm2.Text = Val(frmWbNew.txtGd.Text) + Val(frmWbNew.txtXm.Text)




'打开应收款表
tt = "select * from htFk where htbh='" & frmWbNew.lblHid.Caption & "'"
frmWbNew.adoFk.Close
frmWbNew.adoFk.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWbNew.dtgFk.DataSource = frmWbNew.adoFk
'打开佣金表
tt = "select yED as 收款额度,YingFu as 支付金额,yid,ywy from yongjin where htbh='" & frmWbNew.txtHtbh.Text & "' order by yid"
frmWbNew.adoYj.Close
frmWbNew.adoYj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWbNew.dtgYJ.DataSource = frmWbNew.adoYj
frmWbxjB.frmYm.Visible = False
frmWbNew.tabHt.Tab = 0
frmWbNew.cmdMod.Enabled = True
frmWbNew.cmdSave.Enabled = False
frmWbNew.frmYm.Visible = False
Call modHt.OpenHtAn

frmWbNew.cmdClcb.Visible = False
If mod1.DName = "马晓聪" Or mod1.DName = "倪旭" Then
    frmWbNew.cmdClcb.Visible = True
End If

End Sub

Public Sub NewQing()
frmWbNew.txtKhmc.Text = ""
frmWbNew.txtKhdm.Text = ""
frmWbNew.txtHtbh.Text = ""
frmWbNew.lblHtxz.Caption = ""
frmWbNew.txtXmmc.Text = ""
frmWbNew.txtXmmc.Tag = ""
frmWbNew.txtXYwy.Text = ""
frmWbNew.txtXYwy.Tag = ""
frmWbNew.txtHtrq.Text = ""
frmWbNew.comQy.Text = ""
frmWbNew.txtADR.Text = ""
frmWbNew.txtHtze.Text = ""
frmWbNew.txtZe.Text = ""
frmWbNew.txtEd.Text = ""
Set frmWbNew.dtgFk.DataSource = Nothing
If frmWbNew.dtgFk.Rows > 2 Then frmWbNew.dtgFk.Rows = 1
Set frmWbNew.dtgYf.DataSource = Nothing

frmWbNew.optLa.Value = False
frmWbNew.optLb.Value = False
frmWbNew.optLc.Value = False
frmWbNew.txtCbze1.Text = ""
frmWbNew.txtClcb1.Text = ""
frmWbNew.txtRgf1.Text = ""
frmWbNew.txtCLF1.Text = ""
frmWbNew.txtCLF1.Text = ""
frmWbNew.txtYf1.Text = ""
frmWbNew.txtQt1.Text = ""
frmWbNew.txtJlr1.Text = ""
frmWbNew.txtYj1.Text = ""
frmWbNew.txtLr1.Text = ""
frmWbNew.txtTcBe.Text = ""
frmWbNew.txtTc2.Text = ""
frmWbNew.txtTcRQ.Text = ""
frmWbNew.txtCbze2.Text = ""
frmWbNew.txtFbje2.Text = ""
frmWbNew.txtFbje1.ToolTipText = ""
frmWbNew.txtFbje2.ToolTipText = ""
frmWbNew.txtYf2.Text = ""
frmWbNew.txtQt2.Text = ""
frmWbNew.txtJlr2.Text = ""
frmWbNew.txtYj2.Text = ""
frmWbNew.txtLr2.Text = ""
frmWbNew.txtFbje1.Text = ""
frmWbNew.txtFbje2.Text = ""
frmWbNew.lblHid.Caption = ""
frmWbNew.lblBaoId.Caption = ""
frmWbNew.lblPwf.Caption = ""
frmWbNew.lblLc.Caption = ""
frmWbNew.lblLcRen.Caption = ""
frmWbNew.lblLcUid.Caption = ""
frmWbNew.lblFwid.Caption = ""
frmWbNew.lblNlb.Caption = ""
frmWbNew.lblLcou.Caption = ""
frmWbNew.lblBM.Caption = ""
frmWbNew.lblQy.Caption = ""
frmWbNew.lblYwy.Caption = ""
frmWbNew.lblUid.Caption = ""
frmWbNew.comZu.Text = ""
frmWbNew.txtZu.Text = ""
frmWbNew.comPb.Text = ""
frmWbNew.comXh.Text = ""
frmWbNew.txtSl.Text = ""
frmWbNew.chkBa.Value = 0
frmWbNew.chkBb.Value = 0
frmWbNew.chkBc.Value = 0
frmWbNew.txtMOn.Text = ""
frmWbNew.txtWc.Text = ""
frmWbNew.txtXc.Text = ""
frmWbNew.txtYrq.Text = ""
frmWbNew.txtF.Text = ""
frmWbNew.txtL.Text = ""
frmWbNew.frmPL.Visible = False
frmWbNew.txtTl.Text = ""
frmWbNew.txtDj.Text = ""
frmWbNew.txtBz.Text = ""
Set frmWbNew.dtgWb.DataSource = Nothing
If frmWbNew.dtgWb.Rows > 2 Then frmWbNew.dtgWb.Rows = 1
Set frmWbNew.dtgLj.DataSource = Nothing
If frmWbNew.dtgLj.Rows > 2 Then frmWbNew.dtgLj.Rows = 1
frmWbNew.txtDXNR.Text = ""
Set frmWbNew.dtgBao.DataSource = Nothing
If frmWbNew.dtgBao.Rows > 2 Then frmWbNew.dtgBao.Rows = 1
Set frmWbNew.dtgMa.DataSource = Nothing
If frmWbNew.dtgMa.Rows > 2 Then frmWbNew.dtgMa.Rows = 1
frmWbNew.frmYJ.Visible = False
Set frmWbNew.dtgFk.DataSource = Nothing
If frmWbNew.dtgFk.Rows > 2 Then frmWbNew.dtgFk.Rows = 1

Set frmWbNew.dtgJTf.DataSource = Nothing
If frmWbNew.dtgJTf.Rows > 2 Then frmWbNew.dtgJTf.Rows = 1
Set frmWbNew.dtgQkf.DataSource = Nothing
If frmWbNew.dtgQkf.Rows > 2 Then frmWbNew.dtgQkf.Rows = 1
Set frmWbNew.dtgyjF.DataSource = Nothing
If frmWbNew.dtgyjF.Rows > 2 Then frmWbNew.dtgyjF.Rows = 1
frmWbNew.dtgWb.Visible = False
frmWbNew.dtgLj.Visible = False
frmWbNew.txtDXNR.Visible = False
frmWbNew.dtgBao.Visible = False
frmWbNew.dtgMa.Visible = False
frmWbNew.lblJiLI.Visible = False
frmWbNew.cmdYadd.Visible = False
frmWbNew.cmdYdel.Visible = False
Bid = 0

Set frmWbNew.dtgGD.DataSource = Nothing
frmWbNew.dtgGD.Rows = 1
frmWbNew.optGDA.Value = False
frmWbNew.optGDB.Value = False
frmWbNew.optGDC.Value = False
frmWbNew.txtGDNR.Text = ""
frmWbNew.txtXm.Text = ""
frmWbNew.txtGd.Text = ""
frmWbNew.txtQdj.Text = ""
frmWbNew.txtRl.Text = ""
frmWbNew.cmdKP.Visible = False
frmWbNew.cmdYadd.Visible = False
frmWbNew.cmdDel.Visible = False
frmWbNew.frmJTF.Visible = False
frmWbNew.frmQkF.Visible = False
frmWbNew.txtQkf.Text = ""
frmWbNew.txtYjf.Text = ""
frmWbNew.txtJTf.Text = ""
frmWbNew.frmCw.Enabled = False
frmWbNew.lblHid.Caption = ""
End Sub


Public Sub NewLocked()
frmWbNew.txtKhmc.Locked = True
frmWbNew.txtKhdm.Locked = True

frmWbNew.txtXmmc.Locked = True

frmWbNew.txtXYwy.Locked = True

frmWbNew.txtHtrq.Locked = True
frmWbNew.comQy.Locked = True
frmWbNew.txtADR.Locked = True
frmWbNew.txtHtze.Locked = True
frmWbNew.txtZe.Locked = True
frmWbNew.txtEd.Locked = True

frmWbNew.optLa.Enabled = False
frmWbNew.optLb.Enabled = False
frmWbNew.optLc.Enabled = False
frmWbNew.txtCbze1.Locked = True
frmWbNew.txtClcb1.Locked = True
frmWbNew.txtRgf1.Locked = True
frmWbNew.txtCLF1.Locked = True
frmWbNew.txtCLF1.Locked = True
frmWbNew.txtYf1.Locked = True
frmWbNew.txtQt1.Locked = True
frmWbNew.txtJlr1.Locked = True
frmWbNew.txtYj1.Locked = True
frmWbNew.txtLr1.Locked = True
frmWbNew.txtTcBe.Locked = True
frmWbNew.txtTc2.Locked = True
frmWbNew.txtTcRQ.Locked = True
frmWbNew.txtCbze2.Locked = True
frmWbNew.txtFbje2.Locked = True
frmWbNew.txtYf2.Locked = True
frmWbNew.txtQt2.Locked = True
frmWbNew.txtJlr2.Locked = True
frmWbNew.txtYj2.Locked = True
frmWbNew.txtLr2.Locked = True
frmWbNew.txtFbje1.Locked = True
frmWbNew.txtFbje2.Locked = True
frmWbNew.dt3.Enabled = False
frmWbNew.dt4.Enabled = False
frmWbNew.txtF.Locked = True
frmWbNew.txtL.Locked = True
frmWbNew.frmFX.Visible = False
End Sub

Public Sub HtLcBut(Nlb As Integer)
Dim tt As String
Dim oo As Integer
On Error Resume Next
For oo = 10 To 1 Step -1
    Unload frmWbNew.lblTm(oo)
    Unload frmWbNew.cmdQm(oo)
    Unload frmWbNew.lblQM(oo)
Next
    frmWbNew.cmdQm(0).Caption = ""
    frmWbNew.lblTm(0).Caption = ""
    frmWbNew.cmdQm(0).Visible = True
    frmWbNew.lblQM(0).Visible = True
    frmWbNew.lblTm(0).Visible = True
    tt = "lcBut(" & Nlb & ")"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    mod1.HTP.MoveFirst
    frmWbNew.cmdQm(0).Caption = ""
    frmWbNew.lblQM(0).Caption = mod1.HTP.Fields("LNR").Value
    frmWbNew.lblTm(0).Caption = ""
    mod1.HTP.MoveNext '第一个数组按钮不用添加,所以,跳到下一记录
    For oo = 1 To mod1.HTP.RecordCount - 1
        Load frmWbNew.lblQM(oo)
        Load frmWbNew.cmdQm(oo)
        Load frmWbNew.lblTm(oo)
        frmWbNew.lblQM(oo).Caption = mod1.HTP.Fields("LNR").Value
        frmWbNew.lblQM(oo).Visible = True
        frmWbNew.lblQM(oo).Left = frmWbNew.lblQM(oo - 1).Left + 1200
        frmWbNew.cmdQm(oo).Caption = ""
        frmWbNew.cmdQm(oo).Visible = True
        frmWbNew.cmdQm(oo).Left = frmWbNew.cmdQm(oo - 1).Left + 1200
        frmWbNew.lblTm(oo).Caption = ""
        frmWbNew.lblTm(oo).Visible = True
        frmWbNew.lblTm(oo).Left = frmWbNew.lblTm(oo - 1).Left + 1200
        mod1.HTP.MoveNext
    Next


        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "QMRZAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@NLb") = Nlb
        mod1.cmd.Parameters("@btz") = mod1.BTZ
        mod1.cmd.Parameters("@QDBH") = frmWbNew.txtHtbh.Text
        mod1.cmd.Execute
        Set cmd = Nothing
End Sub

Public Sub OpenHtAn()
Dim tt As String
Dim oo As Integer
On Error Resume Next

    For oo = 10 To 1 Step -1
        Unload frmWbNew.cmdQm(oo)
        Unload frmWbNew.lblQM(oo)
        Unload frmWbNew.lblTm(oo)
    Next
    frmWbNew.cmdQm(0).Caption = ""
    frmWbNew.lblTm(0).Caption = ""
      tt = "qmrzOpen(" & mod1.BTZ & ",'" & frmWbNew.txtHtbh.Text & "')"
      Set mod1.HTP = CreateObject("adodb.recordset")
      mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
      If mod1.HTP.RecordCount > 0 Then
         mod1.HTP.MoveFirst
         frmWbNew.cmdQm(0).Visible = True
         frmWbNew.lblQM(0).Visible = True
         frmWbNew.lblTm(0).Visible = True
                  frmWbNew.lblQM(0).Caption = mod1.HTP.Fields("QLabel").Value
        If mod1.HTP.Fields("xf").Value = True Then
         frmWbNew.cmdQm(0).Caption = mod1.HTP.Fields("Qren").Value
         frmWbNew.lblTm(0).Caption = mod1.HTP.Fields("QRQ").Value
         End If
         frmWbNew.cmdQm(0).Tag = mod1.HTP.Fields("zid").Value
         mod1.HTP.MoveNext
         For oo = 1 To mod1.HTP.RecordCount - 1
           Load frmWbNew.lblQM(oo)
           frmWbNew.lblQM(oo).Caption = ""
           Load frmWbNew.cmdQm(oo)
           frmWbNew.cmdQm(oo).Caption = ""
           Load frmWbNew.lblTm(oo)
           frmWbNew.lblTm(oo).Caption = ""
           frmWbNew.lblQM(oo).Caption = mod1.HTP.Fields("QLabel").Value
            If mod1.HTP.Fields("xf").Value = True Then
                frmWbNew.cmdQm(oo).Caption = mod1.HTP.Fields("Qren").Value
                If frmWbNew.cmdQm(oo).Caption = "南京办经理" Then
                    frmWbNew.cmdQm(oo).Caption = "南京办经理"
                End If
                frmWbNew.lblTm(oo).Caption = mod1.HTP.Fields("QRQ").Value
           End If
           frmWbNew.cmdQm(oo).Tag = mod1.HTP.Fields("zid").Value
           frmWbNew.lblQM(oo).Visible = True
           frmWbNew.cmdQm(oo).Visible = True
           frmWbNew.lblTm(oo).Visible = True
           frmWbNew.lblQM(oo).Left = frmWbNew.lblQM(oo - 1).Left + 1200
           frmWbNew.cmdQm(oo).Left = frmWbNew.cmdQm(oo - 1).Left + 1200
           frmWbNew.lblTm(oo).Left = frmWbNew.lblTm(oo - 1).Left + 1200
           mod1.HTP.MoveNext
        Next
     Else
        frmWbNew.cmdQm(0).Visible = False
        frmWbNew.lblQM(0).Visible = False
        frmWbNew.lblTm(0).Visible = False
     End If

End Sub
