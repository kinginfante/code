Attribute VB_Name = "modNewHT"

Public Sub NewLocked()

FMXC.txtHtze.Locked = True
FMXC.txtBz.Locked = True
FMXC.comFP.Locked = True
FMXC.frmFX.Visible = False
FMXC.txtYf1.Locked = True
FMXC.dt3.Enabled = False '维保起始期
FMXC.dt4.Enabled = False
'FMXC.txtWBNR.Locked = True '外包内容
FMXC.txtTcBe.Locked = True
'FMXC.txtYjf.Locked = True
'FMXC.txtJTf.Locked = True
'FMXC.txtQkf.Locked = True
'FMXC.chkYJF.Enabled = False
'FMXC.chkJTF.Enabled = False
'FMXC.chkQKF.Enabled = False
'FMXC.txtYjfBz.Locked = True
FMXC.frmCw.Enabled = False
End Sub

Public Sub NewMQing()
Dim oo As Integer
On Error Resume Next
FMXC.NewF = 0
For oo = 0 To 5

    FMXC.cmdMQm(oo).Caption = ""
    FMXC.lblMTm(oo).Caption = ""
Next
FMXC.dtgFL.Clear
Call FMXC.FLGG
FMXC.frmDate.Visible = False
FMXC.comKQY.Text = ""
FMXC.comKQY.Locked = True
''''''''FMXC.cmdW1.BackColor = &H8000000F ': FMXC.cmdW1.Visible = False
''''''''FMXC.cmdW2.BackColor = &H8000000F ': FMXC.cmdW2.Visible = False
''''''''FMXC.cmdW3.BackColor = &H8000000F ': FMXC.cmdW3.Visible = False
''''''''FMXC.cmdW4.BackColor = &H8000000F ': FMXC.cmdW4.Visible = False
''''''''FMXC.cmdW5.BackColor = &H8000000F ': FMXC.cmdW5.Visible = False
FMXC.frmFX.Visible = False
FMXC.txtXmmc.Text = ""
FMXC.txtXmmc.ToolTipText = ""  '项目代号
FMXC.txtKhmc.Text = ""
FMXC.txtKhdm.Text = ""
FMXC.txtYwy.Text = ""
FMXC.txtYwy.ToolTipText = "" '业务员代号
FMXC.txtXYwy.Text = ""
FMXC.txtXYwy.ToolTipText = "" '项目管理者代号
FMXC.txtHtbh.Text = ""
FMXC.txtHtrq.Text = ""
FMXC.comQy.Text = ""
FMXC.txtADR.Text = ""
FMXC.cmdWb.ToolTipText = ""
FMXC.txtBz.Text = ""
FMXC.txtZe.Text = "" '实付总额
FMXC.txtEd.Text = "" '实付额度
FMXC.dtgSD.Clear
FMXC.txtHtze.Text = ""
FMXC.comFP.Text = "" '发票情况
FMXC.lblHtxz.Caption = ""
FMXC.txtYrq.Text = ""
FMXC.txtYed.Text = ""
FMXC.txtYje.Text = ""
FMXC.txtYjpw.Text = ""
FMXC.txtYjpw.Visible = False
FMXC.frmFk.Visible = False
FMXC.dtgSD.Visible = True
FMXC.txtH1.Text = ""
'''''FMXC.cmdW1.ToolTipText = ""


FMXC.txtH2.Text = ""
'''''FMXC.cmdW2.ToolTipText = ""

FMXC.txtW3.Text = ""

'''''FMXC.cmdW3.ToolTipText = ""

FMXC.txtW4.Text = ""
'''''FMXC.cmdW4.ToolTipText = ""
FMXC.txtW5.Text = ""
FMXC.txtH5.Text = ""
'''''FMXC.cmdW5.ToolTipText = ""

FMXC.txtW6.Text = ""
FMXC.txtH6.Text = ""
'''''FMXC.cmdW6.ToolTipText = ""

FMXC.txtCbze1.Text = ""
FMXC.txtCbze2.Text = ""
FMXC.txtClcb1.Text = ""
FMXC.txtClcb2.Text = ""
FMXC.txtRgf1.Text = ""
FMXC.txtRGF2.Text = ""
FMXC.txtCLF1.Text = ""
FMXC.txtFbje1.Text = ""
FMXC.txtFbje2.Text = ""
FMXC.txtYf1.Text = ""
FMXC.txtYf2.Text = ""
FMXC.txtQt1.Text = ""
FMXC.txtQt2.Text = ""
FMXC.txtJlr1.Text = ""
FMXC.txtJlr2.Text = ""
FMXC.txtYj1.Text = ""
'FMXC.txtYj2.Text = ""
FMXC.txtLr1.Text = ""
FMXC.txtLr2.Text = ""
FMXC.txtLr2.Visible = False
FMXC.txtTcBe.Text = ""
FMXC.txtTc2.Text = ""
FMXC.txtTcRQ.Text = ""

FMXC.txtF.Text = ""
FMXC.txtL.Text = ""
FMXC.txtWc.Text = ""
FMXC.txtXc.Text = ""
FMXC.comZu.Text = ""
FMXC.txtZu.Text = ""
FMXC.comZuD.Text = ""
FMXC.txtZuD.Text = ""
FMXC.txtDxnr.Text = "" '大修内容
FMXC.txtMon.Text = "" '大修保持期




'FMXC.txtCL.Text = ""
'FMXC.txtCj.Text = ""
'FMXC.txtFbnr.Text = ""
'FMXC.txtWBNR.Text = "" '外包内容


FMXC.MMdtgFk.Clear
FMXC.MMdtgA.Clear
FMXC.MMdtgA.Clear
FMXC.MMdtgBao.Clear
FMXC.MMdtgMa.Clear
FMXC.MMdtgCP.Clear
FMXC.MMdtgCPCB.Clear
FMXC.MMdtgYJ.Clear
FMXC.tabGc.TabVisible(0) = False

FMXC.tabGc.TabVisible(1) = False
FMXC.tabGc.TabVisible(2) = False
FMXC.tabGc.TabVisible(3) = False
FMXC.tabGc.TabVisible(4) = False
FMXC.tabGc.TabVisible(5) = False

FMXC.tabHt.Tab = 0
FMXC.cmdMod.Enabled = True
FMXC.cmdSave.Enabled = False
FMXC.frmYm.Visible = False
FMXC.frmQm.Visible = False

FMXC.txtQM.Text = ""
FMXC.OptT1.Value = True
FMXC.lblTX.Visible = False
FMXC.txtXmmc.Text = "kkk"
FMXC.optY1.Value = False
FMXC.optY2.Value = False

FMXC.chkA.ForeColor = &H80000012
FMXC.chkB.ForeColor = &H80000012
FMXC.chkC.ForeColor = &H80000012
FMXC.chkD.ForeColor = &H80000012
FMXC.chkE.ForeColor = &H80000012
FMXC.chkF.ForeColor = &H80000012
FMXC.cmdYadd.Visible = False
FMXC.cmdYdel.Visible = False
FMXC.lblLc.Caption = ""
FMXC.lblLcRen.Caption = ""
FMXC.lblLcUid.Caption = ""
FMXC.lblyjFF.Caption = ""
FMXC.lblFwid.Caption = ""

FMXC.txtYjf.Text = ""
FMXC.txtJTf.Text = ""
FMXC.txtQkf.Text = ""
FMXC.chkYJF.Value = 0
FMXC.chkJTF.Value = 0
FMXC.chkQKF.Value = 0
FMXC.txtYjfBz.Text = ""
FMXC.txtFC.Text = "" '辅材

FMXC.txtW3.Locked = True
FMXC.txtW4.Locked = True
FMXC.txtW5.Locked = True
FMXC.txtW6.Locked = True
Set mod1.mJt = Nothing
Set mod1.mQk = Nothing
Set FMXC.dtgJTf.DataSource = Nothing
'If FMXC.dtgJTf.Rows > 2 Then FMXC.dtgJTf.Rows = 1
Set FMXC.dtgQkf.DataSource = Nothing
'If FMXC.dtgQkf.Rows > 2 Then FMXC.dtgQkf.Rows = 1
Set FMXC.dtgyjF.DataSource = Nothing
'If FMXC.dtgyjF.Rows > 2 Then FMXC.dtgyjF.Rows = 1
FMXC.frmJTF.Visible = False
FMXC.frmQkF.Visible = False
FMXC.frmCw.Enabled = False
FMXC.MMdtgBao.Visible = False
FMXC.MMdtgMa.Visible = False
FMXC.MMdtgCP.Visible = False
FMXC.MMdtgCPCB.Visible = False

'''''FMXC.cmdW1.Visible = True
'''''FMXC.cmdW2.Visible = True
'''''FMXC.cmdW3.Visible = True
'''''FMXC.cmdW4.Visible = True
'''''FMXC.cmdW5.Visible = True
'''''FMXC.cmdW6.Visible = True
'FMXC.lblFk.Caption = "" '付款方式
FMXC.txtZbh.Text = "" '执行编号
FMXC.txtHtze.Locked = True
FMXC.comYjRen.Text = ""
For oo = 10 To 0 Step -1
    FMXC.comYjRen.RemoveItem oo
Next
FMXC.comYjRen.ToolTipText = ""
FMXC.cmdYview.Caption = ""
FMXC.txtD1.Text = ""
FMXC.txtD2.Text = ""
FMXC.txtD3.Text = ""
FMXC.txtD4.Text = ""
FMXC.txtD5.Text = ""
FMXC.txtD6.Text = ""
FMXC.lblMF.Caption = ""
FMXC.txtHtbh.ToolTipText = ""
End Sub
Public Sub NewB(Hid As Long)
Dim tt As String
Dim oo As Integer
Dim Je As Single
Dim NewF As Integer '合同版本
'版本切换
FMXC.lblWC.Visible = False
FMXC.txtQt1.Visible = False
FMXC.txtW5.Visible = False
FMXC.txtH5.Left = 1770
FMXC.txtH5.Width = 2175
FMXC.txtW6.Visible = False
FMXC.txtH6.Left = 1770
FMXC.txtH6.Width = 2175
FMXC.lblYug.Caption = "基准价"
FMXC.lblYug2.Visible = False
FMXC.chkA.ForeColor = &HC00000
FMXC.chkB.ForeColor = &HC00000
FMXC.chkC.ForeColor = &HC00000
FMXC.chkD.ForeColor = &HC00000
FMXC.chkE.ForeColor = &HC00000
FMXC.chkF.ForeColor = &HC00000
FMXC.txtH1.ForeColor = &HC00000
FMXC.txtH2.ForeColor = &HC00000
FMXC.txtW3.ForeColor = &HC00000
FMXC.txtW4.ForeColor = &HC00000
FMXC.txtH5.ForeColor = &HC00000
FMXC.txtH6.ForeColor = &HC00000
FMXC.lblWC.Visible = False: FMXC.txtQt1.Visible = False
FMXC.lblCBZE.Caption = "基准总价": FMXC.lblCBZE.ForeColor = &HC00000
FMXC.txtCbze1.Width = 2475: FMXC.txtCbze2.Visible = False
FMXC.txtClcb1.Width = 2475: FMXC.txtClcb2.Visible = False
FMXC.txtFbje1.Width = 2475: FMXC.txtFbje2.Visible = False
FMXC.lblCL.Visible = False: FMXC.txtCLF1.Visible = False
FMXC.lblCB.Visible = False
FMXC.lblYug.ForeColor = FMXC.chkA.ForeColor
FMXC.lblClcb.Top = 1650: FMXC.txtClcb1.Top = 1650: FMXC.txtClcb2.Top = 1650
FMXC.lblRG.Top = 2200: FMXC.txtRgf1.Top = 2200
FMXC.txtRGF2.Top = 2200
FMXC.txtRGF2.Visible = False

On Error GoTo ERCU
 Dim ii As Integer
Dim Ra, Rb, RC, RD, RE, Rf, Rg, Rh, Ri, Rj, Rk, RL, RM, RN, RO, RP, RQ
Dim ua, ub, uc, ud, ue, uf, ug, uh, ui, uj, uk, ul, um, un, uo, uq
tt = "declare @bid1 int,@bid2 int,@bid3 int,@bid4 int,@bid5 int,@bid6 int,@rid int,@htbh nvarchar(22);" & _
    "select @bid1=bid1,@bid2=bid2,@bid3=bid3,@bid4=bid4,@bid5=bid5,@bid6=bid6,@htbh=htbh,@rid=rid from htping where hid=" & Hid & ";" & _
 "select khmc,khdh,htbh,htxz,xmmc,xid,Xywy,xuid,ywy,uid,htrq,qy,khadr,htze,bz,fpLX,w11,bid1,w22,bid2,w3,w33,bid3,w4,bid4,w5,w55,bid5,w6,w66,bid6," & _
    "cbze,clCb,rgF,clf1,fbje,yunF,qtF,Jlr1,Yj,rid,xmLr,tcBe,Tc1,TCRQ,yjff,htqy,htqy1,htf,hid,Lc,LcRen,LcUid,Fwid,yjrq,yjf,jtf,qkf,yjbz,fo,fk,zbh,prf,d1,d2,d3,d4,d5,d6,newF,d7,w7,kqy,kqy2,kren,kren2,kuid,kuid2,klb0,klb,klb2 from htping where hid=" & Hid & ";" & _
    "select wc,xc,ta,tb,tc,zname from xunjiaD where bid=@bid1;select jzpb as 机组品牌,jzxh as 机组型号,sl as 数量 from wbjb where bid=@bid1;" & _
    "select zname,mon,dxnr from xunjiaD where bid=@bid2;select jzpb as 机组品牌,jzxh as 机组型号,sl as 数量 from wbjb where bid=@bid2;" & _
    "select dxnr from xunjiaD where bid=@bid3;select dxnr from xunjiaD where bid=@bid4;" & _
     "select * from BaoMxNew where bid=@bid5 order by lid;" & "select * from BaoMxNew where bid=@bid6 order by lid;" & _
     "select rq as 日期,je as 金额,bz as 备注,mid,sum(je) from htpingJt where hid=" & Hid & " and delf=1  group by rq,je,bz,mid  order by mid desc;" & _
     "select rq as 日期,je as 金额,bz as 备注,mid,sum(je) from htpingQk where hid=" & Hid & " and delf=1 group by rq,je,bz,mid  order by mid desc;" & _
     "select 应付日期,收款额度,应付金额,fid from htFK where htbh='" & Hid & "';" & _
     "select yED as 收款额度,YingFu as 支付金额,yid,lc from yongjin where hid=" & Hid & " order by yid;" & _
     "select sum(cjhg) from xunjiaD where lc=100 and htbh=" & Hid & ";" & _
     "select khman from khren where rid=@rid;" & _
     "select fid from hmht.dbo.ht where htbh=@htbh"

     
     
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'Set mod1.HTP = mod1.HTP.NextRecordset
'Set mod1.HTP = mod1.HTP.NextRecordset
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rb = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RC = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RD = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RE = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rf = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rg = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rh = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Ri = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rj = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rk = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RL = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RM = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RN = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RO = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RP = mod1.HTP.GetRows
End If

mod1.HTP.Close
Set mod1.HTP = Nothing

On Error Resume Next

FMXC.Visible = False
FMXC.txtKhmc.Text = Trim(Ra(0, 0))
FMXC.txtKhdm.Text = Trim(Ra(1, 0))
FMXC.txtHtbh.Text = Trim(Ra(2, 0))

FMXC.lblHtxz.Caption = Trim(Ra(3, 0))
FMXC.txtXmmc.Text = Trim(Ra(4, 0))
FMXC.txtXmmc.ToolTipText = Trim(Ra(5, 0))
FMXC.cmdWb.ToolTipText = Trim(Ra(5, 0))
FMXC.txtXYwy.Text = Trim(Ra(6, 0))
FMXC.txtXYwy.ToolTipText = Trim(Ra(7, 0))
FMXC.txtYwy.Text = Trim(Ra(8, 0))
FMXC.txtYwy.ToolTipText = Trim(Ra(9, 0))
FMXC.txtHtrq.Text = Trim(Ra(10, 0))
FMXC.comQy.Text = Trim(Ra(11, 0))
FMXC.txtADR.Text = Trim(Ra(12, 0))
FMXC.txtHtze.Text = Trim(Ra(13, 0))
''''''''''''fmxc.txtZe.Text = mod1.HTP.Fields("htbh").Value  '财务收款(天兴软件)
''''''''''''fmxc.txtEd.Text = mod1.HTP.Fields("htbh").Value

FMXC.txtBz.Text = Trim(Ra(14, 0))
''''''''''fmxc.txtZe.Text = "" '实付总额
''''''''''fmxc.txtEd.Text = "" '实付额度

'开票情况
'''''If Trim(Ra(15, 0)) = "增值发票" Then
'''''    FMXC.optLa.Value = True
'''''ElseIf Trim(Ra(15, 0)) = "商业发票" Then
'''''    FMXC.optLb.Value = True
'''''ElseIf Trim(Ra(15, 0)) = "服务发票" Then
'''''    FMXC.optLc.Value = True
'''''End If
FMXC.comFP.Text = Trim(Ra(15, 0))



FMXC.txtH1.Text = Trim(Ra(16, 0))
If Val(FMXC.txtH1.Text) > 0 Then
    FMXC.chkA.ForeColor = &HC000C0
'''''    FMXC.cmdW1.BackColor = &H8080FF
'''''    FMXC.cmdW1.Visible = True
    FMXC.dtgFL.Col = 2: FMXC.dtgFL.Row = 1: FMXC.dtgFL.Text = Trim(Ra(16, 0))
Else
    '''''If FMXC.txtHtbh.Text <> "HMNEW" Then
        '''''FMXC.cmdW1.Visible = False
    '''''End If
End If
'''''FMXC.cmdW1.ToolTipText = Trim(Ra(17, 0))
FMXC.dtgFL.Col = 4
For oo = 1 To 2
    FMXC.dtgFL.Row = oo
    If FMXC.txtHtbh.Text = "HMNEW" And Val(Ra(17, 0)) = 0 Then
        FMXC.dtgFL.Text = "双击新增"
    Else
        FMXC.dtgFL.Text = "XJD" & Trim(Ra(17, 0))
        If Val(Trim(Ra(17, 0))) = 0 Then FMXC.dtgFL.Text = "豪曼科技"
    End If
Next
'fmxc.chkB.Value = mod1.HTP.Fields("chkB").Value

FMXC.txtH2.Text = Trim(Ra(18, 0))
If Val(FMXC.txtH2.Text) > 0 Then
    FMXC.chkB.ForeColor = &HC000C0
'''''    FMXC.cmdW2.BackColor = &H8080FF
'''''    FMXC.cmdW2.Visible = True
    FMXC.dtgFL.Col = 2: FMXC.dtgFL.Row = 2: FMXC.dtgFL.Text = Trim(Ra(18, 0))
Else
''''''    If FMXC.txtHtbh.Text <> "HMNEW" Then
''''''        FMXC.cmdW2.Visible = False
''''''    End If
End If
'''''FMXC.cmdW2.ToolTipText = Trim(Ra(19, 0))
'fmxc.chkC.Value = mod1.HTP.Fields("chkC").Value
FMXC.txtW3.Text = Trim(Ra(20, 0))
If Val(FMXC.txtW3.Text) > 0 Then
    FMXC.chkC.ForeColor = &HC000C0
'''''    FMXC.cmdW3.BackColor = &H8080FF
'''''    FMXC.cmdW3.Visible = True
    FMXC.dtgFL.Col = 2: FMXC.dtgFL.Row = 3: FMXC.dtgFL.Text = Trim(Ra(20, 0))
Else
'''''    If FMXC.txtHtbh.Text <> "HMNEW" Then
'''''        FMXC.cmdW3.Visible = False
'''''    End If
End If
'FMXC.txtH3.Text = Trim(Ra(21, 0))
'FMXC.cmdW3.ToolTipText = Trim(Ra(22, 0))
If Trim(Ra(22, 0)) > 0 Then
    FMXC.dtgFL.Col = 4: FMXC.dtgFL.Row = 3: FMXC.dtgFL.Text = "XJD" & Trim(Ra(22, 0))
    FMXC.dtgFL.Col = 4: FMXC.dtgFL.Row = 4: FMXC.dtgFL.Text = "XJD" & Trim(Ra(22, 0))
    FMXC.dtgFL.Col = 4: FMXC.dtgFL.Row = 5: FMXC.dtgFL.Text = "XJD" & Trim(Ra(22, 0))
End If
For oo = 3 To 5
    FMXC.dtgFL.Row = oo
    If FMXC.txtHtbh.Text = "HMNEW" And Val(Ra(22, 0)) = 0 Then
        FMXC.dtgFL.Row = 4
        FMXC.dtgFL.Text = "双击新增"
    Else
        FMXC.dtgFL.Text = "XJD" & Trim(Ra(22, 0))
        If Val(Trim(Ra(22, 0))) = 0 Then FMXC.dtgFL.Text = "豪曼科技"
    End If
Next
'fmxc.chkD.Value = mod1.HTP.Fields("chkD").Value
FMXC.txtW4.Text = Trim(Ra(23, 0))
If Val(FMXC.txtW4.Text) > 0 Then
    FMXC.chkD.ForeColor = &HC000C0
''''''    FMXC.cmdW4.BackColor = &H8080FF
''''''    FMXC.cmdW4.Visible = True
    FMXC.dtgFL.Col = 2: FMXC.dtgFL.Row = 4: FMXC.dtgFL.Text = Trim(Ra(23, 0))
Else
'''''    If FMXC.txtHtbh.Text <> "HMNEW" Then
'''''        FMXC.cmdW4.Visible = False
'''''    End If
End If
'''''FMXC.cmdW4.ToolTipText = Trim(Ra(24, 0))
'fmxc.chkE.Value = mod1.HTP.Fields("chkE").Value
FMXC.txtW5.Text = Trim(Ra(25, 0))
FMXC.txtH5.Text = Trim(Ra(26, 0))
If Val(FMXC.txtW5.Text) > 0 Or Val(FMXC.txtH5.Text) > 0 Then
    FMXC.chkE.ForeColor = &HC000C0
'''''    FMXC.cmdW5.BackColor = &H8080FF
'''''    FMXC.cmdW5.Visible = True
    FMXC.dtgFL.Col = 2: FMXC.dtgFL.Row = 6: FMXC.dtgFL.Text = Trim(Ra(26, 0))
Else
'''''    If FMXC.txtHtbh.Text <> "HMNEW" Then
'''''        FMXC.cmdW5.Visible = False
'''''    End If
End If
'''''FMXC.cmdW5.ToolTipText = Trim(Ra(27, 0))
FMXC.dtgFL.Col = 4: FMXC.dtgFL.Row = 6: FMXC.dtgFL.Text = "XJD" & Trim(Ra(27, 0))
    If FMXC.txtHtbh.Text = "HMNEW" And Val(Ra(27, 0)) = 0 Then
        FMXC.dtgFL.Text = "双击新增   "
    End If

'fmxc.chkF.Value = mod1.HTP.Fields("chkF").Value
FMXC.txtW6.Text = Trim(Ra(28, 0))
FMXC.txtH6.Text = Trim(Ra(29, 0))
If Val(FMXC.txtW6.Text) > 0 Or Val(FMXC.txtH6.Text) > 0 Then
    FMXC.chkF.ForeColor = &HC000C0
''''''    FMXC.cmdW6.BackColor = &H8080FF
''''''    FMXC.cmdW6.Visible = True
    FMXC.dtgFL.Col = 2: FMXC.dtgFL.Row = 7: FMXC.dtgFL.Text = Trim(Ra(29, 0))
Else
'''''    If FMXC.txtHtbh.Text <> "HMNEW" Then
'''''        FMXC.cmdW6.Visible = False
'''''    End If
End If
'''''FMXC.cmdW6.ToolTipText = Trim(Ra(30, 0))
FMXC.dtgFL.Col = 4: FMXC.dtgFL.Row = 7: FMXC.dtgFL.Text = "XJD" & Trim(Ra(30, 0))
    If FMXC.txtHtbh.Text = "HMNEW" And Val(Ra(30, 0)) = 0 Then
        FMXC.dtgFL.Text = "双击新增        "
    End If


'''''If Val(FMXC.cmdW1.ToolTipText) > 0 Then
'''''    FMXC.cmdW1.BackColor = &H8080FF
'''''End If
'''''If Val(FMXC.cmdW2.ToolTipText) > 0 Then
'''''    FMXC.cmdW2.BackColor = &H8080FF
'''''End If
'''''If Val(FMXC.cmdW3.ToolTipText) > 0 Then
'''''    FMXC.cmdW3.BackColor = &H8080FF
'''''End If
'''''If Val(FMXC.cmdW4.ToolTipText) > 0 Then
'''''    FMXC.cmdW4.BackColor = &H8080FF
'''''End If
'''''If Val(FMXC.cmdW5.ToolTipText) > 0 Then
'''''    FMXC.cmdW5.BackColor = &H8080FF
'''''End If
'''''If Val(FMXC.cmdW6.ToolTipText) > 0 Then
'''''    FMXC.cmdW6.BackColor = &H8080FF
'''''End If

FMXC.txtCbze1.Text = Trim(Ra(31, 0))

FMXC.txtClcb1.Text = Trim(Ra(32, 0))

FMXC.txtRgf1.Text = Trim(Ra(33, 0))

FMXC.txtCLF1.Text = Trim(Ra(34, 0))
FMXC.txtFbje1.Text = Trim(Ra(35, 0))

FMXC.txtYf1.Text = Trim(Ra(36, 0))

FMXC.txtQt1.Text = Trim(Ra(37, 0))

FMXC.txtJlr1.Text = Trim(Ra(38, 0))
'fmxc.txtJlr2.Text = ""
FMXC.txtYj1.Text = Trim(Ra(39, 0))
'''''FMXC.txtYj2.Text = Trim(Ra(40, 0))
If Val(FMXC.txtYj1.Text) > 0 Then
    FMXC.optY1.Value = True
Else
    FMXC.optY2.Value = True
End If
FMXC.comYjRen.ToolTipText = Val(Ra(40, 0))
FMXC.comYjRen.Text = Trim(RO(0, 0))

FMXC.txtLr1.Text = Trim(Ra(41, 0))
'fmxc.txtLr2.Text = ""

FMXC.txtTcBe.Text = Trim(Ra(42, 0))
FMXC.txtTc2.Text = Trim(Ra(43, 0))
FMXC.txtTcRQ.Text = Trim(Ra(44, 0))
FMXC.lblyjFF.Caption = Trim(Ra(45, 0))


FMXC.txtF.Text = Trim(Ra(46, 0))
FMXC.txtL.Text = Trim(Ra(47, 0))


If Trim(Ra(48, 0)) = 0 Then
    FMXC.optP.Value = True
ElseIf Trim(Ra(48, 0)) = 1 Then
    FMXC.optZ.Value = True
ElseIf Trim(Ra(48, 0)) = 9 Then
    FMXC.optG.Value = True
ElseIf Trim(Ra(48, 0)) = 2 Then
    FMXC.optW.Value = True
End If

FMXC.lblMHid.Caption = Trim(Ra(49, 0))



FMXC.lblLc.Caption = Trim(Ra(50, 0))
If Val(FMXC.lblLc.Caption) = 1 Then
    FMXC.optY1.Value = False
    FMXC.optY2.Value = False
End If
FMXC.lblLcRen.Caption = Trim(Ra(51, 0))
FMXC.lblLcUid.Caption = Trim(Ra(52, 0))
If FMXC.lblLc.Caption = 0 Or FMXC.lblLc.Caption = 1 Then
    FMXC.lblLcRen.Caption = FMXC.txtYwy.Text
    FMXC.lblLcUid.Caption = FMXC.txtYwy.ToolTipText
End If
FMXC.lblFwid.Caption = Trim(Ra(53, 0))



If FMXC.txtHtbh.Text = "HMNEW" And (FMXC.txtYwy.ToolTipText = mod1.DHid Or FMXC.txtXYwy.ToolTipText = mod1.DHid) Then
    FMXC.cmdHT.Visible = True
Else
    FMXC.cmdHT.Visible = False
End If

'财务评定
'FMXC.txtYjf.Text = mod1.HTP.Fields("yjRQ").Value
If Trim(Ra(54, 0)) = "1999-1-1" Then FMXC.txtYjf.Text = ""
If Trim(Ra(55, 0)) = True Then
    FMXC.chkYJF.Value = 1
Else
    FMXC.chkYJF.Value = 0
End If
If Trim(Ra(56, 0)) = True Then
    FMXC.chkJTF.Value = 1
Else
    FMXC.chkJTF.Value = 0
End If
If Trim(Ra(57, 0)) Then
    FMXC.chkQKF.Value = 1
Else
    FMXC.chkQKF.Value = 0
End If
FMXC.txtYjfBz.Text = Trim(Ra(58, 0))
FMXC.FO = Val(Ra(59, 0))
'FMXC.lblFk.Caption = Trim(Ra(60, 0))
FMXC.txtZbh.Text = Trim(Ra(61, 0))
'FMXC.frmPrf.Caption = Ra(62, 0)
'''''''If Val(Ra(62, 0)) = 1 Then '纯配件合同
'''''''    FMXC.cmdW1.Visible = False: FMXC.cmdW2.Visible = False: FMXC.cmdW3.Visible = False: FMXC.cmdW4.Visible = False: FMXC.cmdW6.Visible = False
'''''''ElseIf Val(Ra(62, 0) = 2) Then
'''''''    FMXC.cmdW1.Visible = True: FMXC.cmdW2.Visible = True: FMXC.cmdW3.Visible = True: FMXC.cmdW4.Visible = True: FMXC.cmdW6.Visible = True
'''''''End If

'常驻基准价
FMXC.dtgFL.Col = 2: FMXC.dtgFL.Row = 5: FMXC.dtgFL.Text = Trim(Ra(71, 0)): If FMXC.dtgFL.Text = 0 Then FMXC.dtgFL.Text = ""
FMXC.comKQY.Text = Trim(Ra(72, 0))

'速达金额
FMXC.txtD1.Text = Val(Ra(63, 0))
FMXC.txtD2.Text = Val(Ra(64, 0))
FMXC.txtD3.Text = Val(Ra(65, 0))
FMXC.txtD4.Text = Val(Ra(66, 0))
FMXC.txtD5.Text = Val(Ra(67, 0))
FMXC.txtD6.Text = Val(Ra(68, 0))
FMXC.dtgFL.Col = 3: FMXC.dtgFL.Row = 1: FMXC.dtgFL.Text = Trim(Ra(63, 0)): If FMXC.dtgFL.Text = 0 Then FMXC.dtgFL.Text = ""
FMXC.dtgFL.Col = 3: FMXC.dtgFL.Row = 2: FMXC.dtgFL.Text = Trim(Ra(64, 0)): If FMXC.dtgFL.Text = 0 Then FMXC.dtgFL.Text = ""
FMXC.dtgFL.Col = 3: FMXC.dtgFL.Row = 3: FMXC.dtgFL.Text = Trim(Ra(65, 0)): If FMXC.dtgFL.Text = 0 Then FMXC.dtgFL.Text = ""
FMXC.dtgFL.Col = 3: FMXC.dtgFL.Row = 4: FMXC.dtgFL.Text = Trim(Ra(66, 0)): If FMXC.dtgFL.Text = 0 Then FMXC.dtgFL.Text = ""
FMXC.dtgFL.Col = 3: FMXC.dtgFL.Row = 5: FMXC.dtgFL.Text = Trim(Ra(70, 0)): If FMXC.dtgFL.Text = 0 Then FMXC.dtgFL.Text = ""
FMXC.dtgFL.Col = 3: FMXC.dtgFL.Row = 6: FMXC.dtgFL.Text = Trim(Ra(67, 0)): If FMXC.dtgFL.Text = 0 Then FMXC.dtgFL.Text = ""
FMXC.dtgFL.Col = 3: FMXC.dtgFL.Row = 7: FMXC.dtgFL.Text = Trim(Ra(68, 0)): If FMXC.dtgFL.Text = 0 Then FMXC.dtgFL.Text = ""

'合同版本
NewF = Val(Ra(69, 0))
FMXC.NewF = NewF
Call FMXC.FLGG
If NewF = 5 Then
    FMXC.dtgFL.Visible = True
'''    FMXC.dtgFL.RowHeight(3) = 0
'''    FMXC.dtgFL.RowHeight(4) = 0
'''    FMXC.dtgFL.RowHeight(5) = 0
    FMXC.dtgFL.MergeCol(4) = True
    FMXC.dtgFL.MergeCells = flexMergeRestrictColumns
ElseIf NewF = 7 Then
    FMXC.dtgFL.Visible = True
'''    FMXC.dtgFL.RowHeight(3) = 0
'''    FMXC.dtgFL.RowHeight(4) = 0
'''    FMXC.dtgFL.RowHeight(5) = 0
    FMXC.dtgFL.MergeCol(4) = True
    FMXC.dtgFL.MergeCells = flexMergeRestrictColumns
Else
    FMXC.dtgFL.Visible = False
End If

'维保
If Val(FMXC.txtH1.Text) > 0 Then
    FMXC.txtWc.Text = Trim(Rb(0, 0))
    FMXC.txtXc.Text = Trim(Rb(1, 0))
    FMXC.chkBa.Value = Trim(Rb(2, 0))
    FMXC.chkBb.Value = Trim(Rb(3, 0))
    FMXC.chkBc.Value = Trim(Rb(4, 0))
    FMXC.txtZu.Text = Trim(Rb(5, 0))
    
    uc = UBound(RC, 2)
    FMXC.MMdtgA.Clear
    FMXC.MMdtgA.Row = 0
    FMXC.MMdtgA.Col = 1
    FMXC.MMdtgA.Text = "机组品牌"
    FMXC.MMdtgA.Col = 2
    FMXC.MMdtgA.Text = "机组型号"
    FMXC.MMdtgA.Col = 3
    FMXC.MMdtgA.Text = "机组数量"
    For oo = 1 To uc + 1
        FMXC.MMdtgA.Row = oo
        For ii = 1 To 3
            FMXC.MMdtgA.Col = ii
            FMXC.MMdtgA.Text = RC(ii, oo)
        Next
    Next
    FMXC.tabGc.TabVisible(0) = True
End If
'大修
If Val(FMXC.txtH2.Text) > 0 Then
    FMXC.txtZuD.Text = RD(0, 0)
    FMXC.txtMon.Text = RD(1, 0)
    FMXC.txtDxnr.Text = RD(2, 0)

    ud = UBound(RD, 2)
    FMXC.MMdtgB.Clear
    FMXC.MMdtgB.Row = 0
    FMXC.MMdtgB.Col = 1
    FMXC.MMdtgB.Text = "机组品牌"
    FMXC.MMdtgB.Col = 2
    FMXC.MMdtgB.Text = "机组型号"
    FMXC.MMdtgB.Col = 3
    FMXC.MMdtgB.Text = "机组数量"
    For oo = 1 To ud + 1
        FMXC.MMdtgB.Row = oo
        For ii = 1 To 3
            FMXC.MMdtgB.Col = ii
            FMXC.MMdtgB.Text = RE(ii, oo)
        Next
    Next
    FMXC.tabGc.TabVisible(1) = True
End If

'If fmxc.chkC.Value = 1 Then '工程分包
If Val(FMXC.txtW3.Text) > 0 Then '工程分包
    'FMXC.txtFbnr.Text = Rf(0, 0)
    FMXC.tabGc.TabVisible(4) = True
End If

'If fmxc.chkD.Value = 1 Then '外包
If Val(FMXC.txtW4.Text) > 0 Then '
    FMXC.txtDxnr.Text = Rg(0, 0)
    FMXC.tabGc.TabVisible(5) = True
End If


FMXC.MMdtgBao.Visible = False
FMXC.MMdtgMa.Visible = False
FMXC.MMdtgCP.Visible = False
FMXC.MMdtgCPCB.Visible = False


FMXC.cmdGx.Visible = False
'FMXC.Label54.Visible = False
'FMXC.txtCj.Visible = False
'FMXC.cmdCGX.Visible = False
If Val(FMXC.txtH5.Text) > 0 Then '配件
    FMXC.MMdtgBao.FixedRows = 1
    uh = UBound(Rh, 2)
    FMXC.MMdtgBao.Row = 0: FMXC.MMdtgBao.Col = 1
    FMXC.MMdtgBao.Text = "品种"
    FMXC.MMdtgBao.Col = 2: FMXC.MMdtgBao.Text = "机组品牌"
    FMXC.MMdtgBao.Col = 3: FMXC.MMdtgBao.Text = "机组型号"
    FMXC.MMdtgBao.Col = 4: FMXC.MMdtgBao.Text = "压缩机型号"
    FMXC.MMdtgBao.Col = 5: FMXC.MMdtgBao.Text = "出厂编号"
    FMXC.MMdtgBao.Col = 6: FMXC.MMdtgBao.Text = "机组序列号"
    FMXC.MMdtgBao.Col = 7: FMXC.MMdtgBao.Text = "零件编号"
    FMXC.MMdtgBao.Col = 8: FMXC.MMdtgBao.Text = "零件名称"
    FMXC.MMdtgBao.Col = 9: FMXC.MMdtgBao.Text = "品牌产地"
    FMXC.MMdtgBao.Col = 10: FMXC.MMdtgBao.Text = "到货期"
    FMXC.MMdtgBao.Col = 11: FMXC.MMdtgBao.Text = "数量"
    FMXC.MMdtgBao.Col = 12: FMXC.MMdtgBao.Text = "单价"
    FMXC.MMdtgBao.Col = 13: FMXC.MMdtgBao.Text = "合计"
    FMXC.MMdtgBao.Col = 14: FMXC.MMdtgBao.Text = "报价有效期"
    FMXC.MMdtgBao.Col = 15: FMXC.MMdtgBao.Text = "baoid"
    FMXC.MMdtgBao.Col = 16: FMXC.MMdtgBao.Text = "Lid"
    FMXC.MMdtgBao.Col = 17: FMXC.MMdtgBao.Text = "Llid"
    FMXC.MMdtgBao.Col = 18: FMXC.MMdtgBao.Text = "bid"
    '显示配件列表
    For oo = 1 To uh + 1
        FMXC.MMdtgBao.Row = oo
        For ii = 1 To 18
            FMXC.MMdtgBao.Col = ii
            FMXC.MMdtgBao.Text = Trim(Rh(ii - 1, oo - 1))
        Next
    Next
   
    
    FMXC.MMdtgBao.MergeCol(1) = True
    FMXC.MMdtgBao.MergeCol(2) = True
    FMXC.MMdtgBao.MergeCol(10) = True
    FMXC.MMdtgBao.MergeCol(14) = True
    FMXC.MMdtgBao.MergeCells = 3

    '显示成本表
     FMXC.MMdtgMa.Row = 0: FMXC.MMdtgMa.Col = 1: FMXC.MMdtgMa.Text = "数量"
     FMXC.MMdtgMa.Col = 2: FMXC.MMdtgMa.Text = "外包单价"
     FMXC.MMdtgMa.Col = 3: FMXC.MMdtgMa.Text = "基准单价"
     FMXC.MMdtgMa.Col = 4: FMXC.MMdtgMa.Text = "外包合计"
     FMXC.MMdtgMa.Col = 5: FMXC.MMdtgMa.Text = "基准合计"

    FMXC.tabGc.TabVisible(2) = True
    FMXC.MMdtgBao.Visible = True
    FMXC.MMdtgMa.Visible = True


    FMXC.cmdGx.Visible = True
End If

'If fmxc.chkF.Value = 1 Then '产品
If Val(FMXC.txtH6.Text) > 0 Then '产品
    FMXC.MMdtgCP.FixedRows = 1
    '显示产品列表
    ui = UBound(Ri, 2)
    FMXC.MMdtgCP.Row = 0: FMXC.MMdtgCP.Col = 1
    FMXC.MMdtgCP.Text = "品种"
    FMXC.MMdtgCP.Col = 2: FMXC.MMdtgCP.Text = "机组品牌"
    FMXC.MMdtgCP.Col = 3: FMXC.MMdtgCP.Text = "机组型号"
    FMXC.MMdtgCP.Col = 4: FMXC.MMdtgCP.Text = "压缩机型号"
    FMXC.MMdtgCP.Col = 5: FMXC.MMdtgCP.Text = "出厂编号"
    FMXC.MMdtgCP.Col = 6: FMXC.MMdtgCP.Text = "机组序列号"
    FMXC.MMdtgCP.Col = 7: FMXC.MMdtgCP.Text = "零件编号"
    FMXC.MMdtgCP.Col = 8: FMXC.MMdtgCP.Text = "零件名称"
    FMXC.MMdtgCP.Col = 9: FMXC.MMdtgCP.Text = "品牌产地"
    FMXC.MMdtgCP.Col = 10: FMXC.MMdtgCP.Text = "到货期"
    FMXC.MMdtgCP.Col = 11: FMXC.MMdtgCP.Text = "数量"
    FMXC.MMdtgCP.Col = 12: FMXC.MMdtgCP.Text = "单价"
    FMXC.MMdtgCP.Col = 13: FMXC.MMdtgCP.Text = "合计"
    FMXC.MMdtgCP.Col = 14: FMXC.MMdtgCP.Text = "报价有效期"
    FMXC.MMdtgCP.Col = 15: FMXC.MMdtgCP.Text = "baoid"
    FMXC.MMdtgCP.Col = 16: FMXC.MMdtgCP.Text = "Lid"
    FMXC.MMdtgCP.Col = 17: FMXC.MMdtgCP.Text = "Llid"
    FMXC.MMdtgCP.Col = 18: FMXC.MMdtgCP.Text = "bid"
    
    '显示配件列表
    For oo = 1 To ui + 1
        FMXC.MMdtgCP.Row = oo
        For ii = 1 To 18
            FMXC.MMdtgCP.Col = ii
            FMXC.MMdtgCP.Text = Trim(Ri(ii - 1, oo - 1))
        Next
    Next
    
    FMXC.MMdtgCP.MergeCol(1) = True
    FMXC.MMdtgCP.MergeCol(2) = True
    FMXC.MMdtgCP.MergeCol(10) = True
    FMXC.MMdtgCP.MergeCol(14) = True
    FMXC.MMdtgCP.MergeCells = 3

    '显示成本表
     FMXC.MMdtgCPCB.Row = 0: FMXC.MMdtgCPCB.Col = 1: FMXC.MMdtgCPCB.Text = "数量"
     FMXC.MMdtgCPCB.Col = 2: FMXC.MMdtgCPCB.Text = "外包单价"
     FMXC.MMdtgCPCB.Col = 3: FMXC.MMdtgCPCB.Text = "基准单价"
     FMXC.MMdtgCPCB.Col = 4: FMXC.MMdtgCPCB.Text = "外包合计"
     FMXC.MMdtgCPCB.Col = 5: FMXC.MMdtgCPCB.Text = "基准合计"
    If mod1.Bm <> "商务部" Then
        FMXC.MMdtgCPCB.ColWidth(1) = 0: FMXC.MMdtgCPCB.ColWidth(3) = 0
    End If

    FMXC.tabGc.TabVisible(3) = True
    FMXC.MMdtgCP.Visible = True
    FMXC.MMdtgCPCB.Visible = True
   ' FMXC.Label54.Visible = True
    'FMXC.txtCj.Visible = True
    'FMXC.cmdCGX.Visible = True
End If



'提成
    FMXC.dtgJTf.Row = 0: FMXC.dtgJTf.Col = 1: FMXC.dtgJTf.Text = "日期"
    FMXC.dtgJTf.Col = 2: FMXC.dtgJTf.Text = "金额"
    FMXC.dtgJTf.Col = 3: FMXC.dtgJTf.Text = "备注"
    FMXC.dtgJTf.Col = 4: FMXC.dtgJTf.Text = "mid"
    For oo = 1 To uj + 1
        FMXC.dtgJTf.Row = oo
        For ii = 1 To 4
            FMXC.dtgJTf.Col = ii
            FMXC.dtgJTf.Text = Trim(Rj(ii - 1, oo - 1))
        Next
    Next
    FMXC.txtJTf.Text = Rj(5, 0)


'业绩
    FMXC.dtgyjF.Row = 0: FMXC.dtgyjF.Col = 1: FMXC.dtgyjF.Text = "日期"
    FMXC.dtgyjF.Col = 2: FMXC.dtgyjF.Text = "金额"
    FMXC.dtgyjF.Col = 3: FMXC.dtgyjF.Text = "备注"
    FMXC.dtgyjF.Col = 4: FMXC.dtgyjF.Text = "mid"
    For oo = 1 To uk + 1
        FMXC.dtgyjF.Row = oo
        For ii = 1 To 4
            FMXC.dtgyjF.Col = ii
            FMXC.dtgyjF.Text = Rj(ii - 1, oo - 1)
        Next
    Next
    FMXC.txtYjf.Text = Rk(5, 0)


''''''''''tt = "SELECT rp_dd as 日期,amtn_cls as 金额,rem as 备注 FROM TF_MON where rp_id=1 and cas_no='" & FMXC.txtHtbh.Text & "' order by rp_dd"
''''''''''
''''''''''mod1.mQk.Close
''''''''''mod1.mQk.Open tt, mod1.workTx, adOpenKeyset, adLockReadOnly, adCmdText
''''''''''If IsNull(mod1.mQk.RecordCount) = True Then
''''''''''    MsgBox ("读取数据错误2.5!")
''''''''''    Exit Sub
''''''''''End If
''''''''''
''''''''''If mod1.mQk.RecordCount = 0 Then
''''''''''    Set FMXC.dtgQkf.DataSource = mod1.mQk
''''''''''    FMXC.dtgQkf.Rows = 2
''''''''''    FMXC.dtgQkf.FixedRows = 0
''''''''''    FMXC.dtgQkf.FixedRows = 1
''''''''''
''''''''''Else
''''''''''    FMXC.dtgQkf.Rows = 2
''''''''''    FMXC.dtgQkf.FixedRows = 1
''''''''''    Set FMXC.dtgQkf.DataSource = mod1.mQk
''''''''''End If
''''''''''
'''''''''''tt = "select sum(je) as je from htpingQk where hid=" & Val(FMXC.lblMHid.Caption) & " and delf=1"
''''''''''tt = "select sum(amtn_cls) as je from tf_mon where rp_id=1 and cas_no='" & FMXC.txtHtbh.Text & "'"
''''''''''Set mod1.HTP = CreateObject("adodb.recordset")
''''''''''mod1.HTP.Open tt, mod1.workTx, adOpenKeyset, adLockReadOnly, adCmdText
''''''''''If IsNull(mod1.HTP.RecordCount) = True Then
''''''''''    MsgBox ("读取数据错误2.6!")
''''''''''    Exit Sub
''''''''''End If
''''''''''FMXC.txtQkf.Text = mod1.HTP.Fields("je").Value
''''''''''FMXC.txtZe.Text = FMXC.txtQkf.Text
''''''''''FMXC.txtEd.Text = Round(Val(FMXC.txtZe.Text) / Val(FMXC.txtHtze.Text) * 100, 2)




'打开应收款表
ul = UBound(RL, 2)
FMXC.MMdtgFk.Row = 0: FMXC.MMdtgFk.Col = 1: FMXC.MMdtgFk.Text = "应付日期"
FMXC.MMdtgFk.Col = 2: FMXC.MMdtgFk.Text = "收款额度"
FMXC.MMdtgFk.Col = 3: FMXC.MMdtgFk.Text = "应付金额"
FMXC.MMdtgFk.Rows = 30
For oo = 1 To ul + 1
    FMXC.MMdtgFk.Row = oo
    For ii = 1 To 4
        FMXC.MMdtgFk.Col = ii
        FMXC.MMdtgFk.Text = Trim(RL(ii - 1, oo - 1))
        If ii = 2 Then
            FMXC.MMdtgFk.Text = Str(Val(FMXC.MMdtgFk.Text) * 100) & "%"
        End If
    Next
Next




'打开佣金表
um = UBound(RM, 2)
FMXC.MMdtgYJ.Row = 0: FMXC.MMdtgYJ.Col = 1: FMXC.MMdtgYJ.Text = "收款额度"
FMXC.MMdtgYJ.Col = 2: FMXC.MMdtgYJ.Text = "支付金额"
For oo = 1 To um + 1
    FMXC.MMdtgYJ.Row = oo
    For ii = 1 To 4
        FMXC.MMdtgYJ.Col = ii
        FMXC.MMdtgYJ.Text = Trim(RM(ii - 1, oo - 1))
    Next
Next

Dim CB As Double
FMXC.MMdtgYJ.Row = 0: FMXC.MMdtgYJ.Col = 5: FMXC.MMdtgYJ.Text = "参考额度"
FMXC.MMdtgYJ.Row = 1
FMXC.MMdtgYJ.Col = 1
Do While Not Val(FMXC.MMdtgYJ.Text) = 0
    
    CB = (Val(FMXC.txtHtze.Text) - Val(FMXC.txtCbze1.Text)) * Val(FMXC.MMdtgYJ.Text)
    FMXC.MMdtgYJ.Col = 5
    FMXC.MMdtgYJ.Text = CB
    FMXC.MMdtgYJ.Col = 1
    FMXC.MMdtgYJ.Row = FMXC.MMdtgYJ.Row + 1
    CB = 0
Loop
'显示辅材
FMXC.txtFC.Text = RN(0, 0)

Call modNewHT.OAn


FMXC.frmYm.Visible = False
FMXC.frmYj.Visible = False
FMXC.Visible = True


FMXC.txtHtbh.ToolTipText = RP(0, 0) '电子合同



If FMXC.optZ.Value = True Then
    FMXC.cmdMod.Enabled = False
    FMXC.cmdDel.Enabled = False
    FMXC.cmdSave.Enabled = False
End If
If mod1.DName = "马晓聪" Or mod1.DName = "乔继敏" Then
    FMXC.cmdMod.Enabled = True
    FMXC.cmdDel.Enabled = False
    FMXC.cmdSave.Enabled = False
End If

If (FMXC.optZ.Value = True Or FMXC.optW.Value = True) And FMXC.txtXYwy.ToolTipText = mod1.DHid Then
    FMXC.cmdNew.Visible = True
Else
    FMXC.cmdNew.Visible = False
End If
If mod1.DName = "谢雪梅" Or FMXC.txtHtrq >= #4/1/2009# Then
    FMXC.dtgFL.Visible = True
Else
    FMXC.dtgFL.Visible = False
End If
    FMXC.cmdDel.Enabled = False
    FMXC.cmdSave.Enabled = False
    FMXC.Show
    FMXC.ZOrder 0
    FMXC.lblMF.Caption = "MF: " & Round((Val(FMXC.txtHtze.Text) - Val(FMXC.txtYj1.Text)) / Val(FMXC.txtCbze1.Text), 2)
If FMXC.txtHtbh.ToolTipText = "" And Val(FMXC.lblLc.Caption) = 1 And FMXC.lblLcUid.Caption = mod1.DName And FMXC.txtHtbh.Text <> "HMNEW" Then
    MsgBox "双击合同编号,可以附加电子合同"
End If
    FMXC.dtgSD.Row = 0
    FMXC.dtgSD.Col = 0: FMXC.dtgSD.Text = "速达入帐日期"
    FMXC.dtgSD.Col = 1: FMXC.dtgSD.Text = "金额"
If FMXC.optP.Value = False Then
    tt = "select billdate,amount from SDV_ChargeA where htbh='" & FMXC.txtHtbh.Text & "' order by billdate"
    Set mod1.HTP = CreateObject("adodb.recordset")
    On Error GoTo ERCU
    On Error Resume Next
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    'On Error Resume Next
    If mod1.HTP.BOF = False Then
        RQ = mod1.HTP.GetRows
    End If
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    '速达入帐
    uq = UBound(RQ, 2) + 1
    FMXC.dtgSD.Rows = uq + 20
    For oo = 1 To uq
        FMXC.dtgSD.Row = oo
        For ii = 0 To 1
            FMXC.dtgSD.Col = ii
            FMXC.dtgSD.Text = RQ(ii, oo - 1)
            If ii = 1 Then
                Je = Val(FMXC.dtgSD.Text) + Je
            End If
        Next
    Next


    FMXC.txtZe.Text = Je
    FMXC.txtEd.Text = Str(Round(Val(FMXC.txtZe.Text) / Val(FMXC.txtHtze.Text), 2) * 100)
End If
Call FmxcFK.Qing
'kqy2,kren,kren2,kuid,kuid2,klb0,klb,klb2
FmxcFK.comQy2.Text = Trim(Ra(72, 0))
FmxcFK.comQy3.Text = Trim(Ra(73, 0))
FmxcFK.txtRen2.Text = Trim(Ra(74, 0))
FmxcFK.txtRen3.Text = Trim(Ra(75, 0))
FmxcFK.txtRen2.ToolTipText = Trim(Ra(76, 0))
FmxcFK.txtRen3.ToolTipText = Trim(Ra(77, 0))
FmxcFK.txtBL1.Text = Trim(Ra(78, 0))
FmxcFK.txtBL2.Text = Trim(Ra(79, 0))
FmxcFK.txtBL3.Text = Trim(Ra(80, 0))

If FMXC.lblHtxz.Caption = "维保" Or FMXC.txtHtbh.Text = "HMNEW" Then
    FMXC.frmDate.Visible = True
End If
Exit Sub
ERCU:
MsgBox ("出错")
End
End Sub


Public Sub NewMBound(Hid As Long)
Dim tt As String
Dim Je As Single
Dim oo As Integer
Dim ii As Integer
'版本切换
FMXC.lblWC.Visible = True
FMXC.txtQt1.Visible = True
FMXC.txtW5.Visible = True
FMXC.txtH5.Left = 3120
FMXC.txtH5.Width = 915
FMXC.txtW6.Visible = True
FMXC.txtH6.Left = 3120
FMXC.txtH6.Width = 915
FMXC.lblYug.Caption = "预估成本"
FMXC.lblYug2.Visible = True
FMXC.dtgSD.Visible = True
FMXC.chkA.ForeColor = &H80000012
FMXC.chkB.ForeColor = &H80000012
FMXC.chkC.ForeColor = &H80000012
FMXC.chkD.ForeColor = &H80000012
FMXC.chkE.ForeColor = &H80000012
FMXC.chkF.ForeColor = &H80000012
FMXC.txtH1.ForeColor = &H80000012
FMXC.txtH2.ForeColor = &H80000012
FMXC.txtW3.ForeColor = &H80000012
FMXC.txtW4.ForeColor = &H80000012
FMXC.txtH5.ForeColor = &H80000012
FMXC.txtH6.ForeColor = &H80000012
FMXC.lblWC.Visible = True: FMXC.txtQt1.Visible = True
FMXC.lblCBZE.Caption = "成本总额": FMXC.lblCBZE.ForeColor = &H80000012
FMXC.txtCbze1.Width = 1245: FMXC.txtCbze2.Visible = True
FMXC.txtClcb1.Width = 1245: FMXC.txtClcb2.Visible = True
FMXC.txtFbje1.Width = 1245: FMXC.txtFbje2.Visible = True
FMXC.lblCL.Visible = True: FMXC.txtCLF1.Visible = True
FMXC.lblCB.Visible = True
FMXC.lblYug.ForeColor = FMXC.chkA.ForeColor
FMXC.lblClcb.Top = 1410: FMXC.txtClcb1.Top = 1410: FMXC.txtClcb2.Top = 1410
FMXC.lblRG.Top = 1875: FMXC.txtRgf1.Top = 1875
FMXC.txtRGF2.Top = 1875

FMXC.dtgFL.Visible = False


On Error GoTo ERCU

Dim Ra, Rb, RC, RD, RE, Rf, Rg, Rh, Ri, Rj, Rk, RL, RM, RP, RQ
Dim ua, ub, uc, ud, ue, uf, ug, uh, ui, uj, uk, ul, um, uq
tt = "declare @bid1 int,@bid2 int,@bid3 int,@bid4 int,@bid5 int,@bid6 int,@htbh nvarchar(22);" & _
    "select @bid1=bid1,@bid2=bid2,@bid3=bid3,@bid4=bid4,@bid5=bid5,@bid6=bid6,@htbh=htbh from htping where hid=" & Hid & ";" & _
 "select khmc,khdh,htbh,htxz,xmmc,xid,Xywy,xuid,ywy,uid,htrq,qy,khadr,htze,bz,fpLX,w11,bid1,w22,bid2,w3,w33,bid3,w4,bid4,w5,w55,bid5,w6,w66,bid6," & _
    "cbze,clCb,rgF,clf1,fbje,yunF,qtF,Jlr1,Yj,Yj1,xmLr,tcBe,Tc1,TCRQ,yjff,htqy,htqy1,htf,hid,Lc,LcRen,LcUid,Fwid,yjrq,yjf,jtf,qkf,yjbz,fo,fk,zbh,prf,d1,d2,d3,d4,d5,d6,kqy,kqy2,kren,kren2,kuid,kuid2,klb0,klb,klb2 from htping where hid=" & Hid & ";" & _
    "select wc,xc,ta,tb,tc,zname from xunjiaD where bid=@bid1;select jzpb as 机组品牌,jzxh as 机组型号,sl as 数量 from wbjb where bid=@bid1;" & _
    "select zname,mon,dxnr from xunjiaD where bid=@bid2;select jzpb as 机组品牌,jzxh as 机组型号,sl as 数量 from wbjb where bid=@bid2;" & _
    "select dxnr from xunjiaD where bid=@bid3;select dxnr from xunjiaD where bid=@bid4;" & _
     "select * from BaoMxNew where bid=@bid5 order by lid;" & "select * from BaoMxNew where bid=@bid6 order by lid;" & _
     "select rq as 日期,je as 金额,bz as 备注,mid,sum(je) from htpingJt where hid=" & Hid & " and delf=1  group by rq,je,bz,mid  order by mid desc;" & _
     "select rq as 日期,je as 金额,bz as 备注,mid,sum(je) from htpingQk where hid=" & Hid & " and delf=1 group by rq,je,bz,mid  order by mid desc;" & _
     "select 应付日期,收款额度,应付金额,fid from htFK where htbh='" & Hid & "';" & _
     "select yED as 收款额度,YingFu as 支付金额,yid from yongjin where htbh=@htbh order by yid;" & _
     "select fid from hmht.dbo.ht where htbh=@htbh;" & _
      "select billdate,amount from SDV_ChargeA where htbh=@htbh order by billdate"
     
     
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'Set mod1.HTP = mod1.HTP.NextRecordset
'Set mod1.HTP = mod1.HTP.NextRecordset
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rb = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RC = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RD = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RE = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rf = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rg = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rh = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Ri = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rj = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rk = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RL = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RM = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RP = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RQ = mod1.HTP.GetRows
End If
mod1.HTP.Close
Set mod1.HTP = Nothing

On Error Resume Next

FMXC.Visible = False
FMXC.txtKhmc.Text = Trim(Ra(0, 0))
FMXC.txtKhdm.Text = Trim(Ra(1, 0))
FMXC.txtHtbh.Text = Trim(Ra(2, 0))

FMXC.lblHtxz.Caption = Trim(Ra(3, 0))
FMXC.txtXmmc.Text = Trim(Ra(4, 0))
FMXC.txtXmmc.ToolTipText = Trim(Ra(5, 0))
FMXC.cmdWb.ToolTipText = Trim(Ra(5, 0))
FMXC.txtXYwy.Text = Trim(Ra(6, 0))
FMXC.txtXYwy.ToolTipText = Trim(Ra(7, 0))
FMXC.txtYwy.Text = Trim(Ra(8, 0))
FMXC.txtYwy.ToolTipText = Trim(Ra(9, 0))
FMXC.txtHtrq.Text = Trim(Ra(10, 0))
FMXC.comQy.Text = Trim(Ra(11, 0))
FMXC.txtADR.Text = Trim(Ra(12, 0))
FMXC.txtHtze.Text = Trim(Ra(13, 0))
''''''''''''fmxc.txtZe.Text = mod1.HTP.Fields("htbh").Value  '财务收款(天兴软件)
''''''''''''fmxc.txtEd.Text = mod1.HTP.Fields("htbh").Value

FMXC.txtBz.Text = Trim(Ra(14, 0))
''''''''''fmxc.txtZe.Text = "" '实付总额
''''''''''fmxc.txtEd.Text = "" '实付额度

'开票情况
'''''If Trim(Ra(15, 0)) = "增值发票" Then
'''''    FMXC.optLa.Value = True
'''''ElseIf Trim(Ra(15, 0)) = "商业发票" Then
'''''    FMXC.optLb.Value = True
'''''ElseIf Trim(Ra(15, 0)) = "服务发票" Then
'''''    FMXC.optLc.Value = True
'''''End If
FMXC.comFP.Text = Trim(Ra(15, 0))


FMXC.txtH1.Text = Trim(Ra(16, 0))

If Val(FMXC.txtH1.Text) > 0 Then
    FMXC.chkA.ForeColor = &HC000C0
''''''    FMXC.cmdW1.BackColor = &H8080FF
''''''    FMXC.cmdW1.Visible = True
Else
'''    If FMXC.txtHtbh.Text <> "HMNEW" Then
'''        FMXC.cmdW1.Visible = False
'''    End If
End If

'fmxc.chkB.Value = mod1.HTP.Fields("chkB").Value

FMXC.txtH2.Text = Trim(Ra(18, 0))
If Val(FMXC.txtH2.Text) > 0 Then
    FMXC.chkB.ForeColor = &HC000C0
End If

'fmxc.chkC.Value = mod1.HTP.Fields("chkC").Value
FMXC.txtW3.Text = Trim(Ra(20, 0))
If Val(FMXC.txtW3.Text) > 0 Then
    FMXC.chkC.ForeColor = &HC000C0

End If
'FMXC.txtH3.Text = Trim(Ra(21, 0))

'fmxc.chkD.Value = mod1.HTP.Fields("chkD").Value
FMXC.txtW4.Text = Trim(Ra(23, 0))
If Val(FMXC.txtW4.Text) > 0 Then
    FMXC.chkD.ForeColor = &HC000C0


End If

'fmxc.chkE.Value = mod1.HTP.Fields("chkE").Value
FMXC.txtW5.Text = Trim(Ra(25, 0))
FMXC.txtH5.Text = Trim(Ra(26, 0))
If Val(FMXC.txtW5.Text) > 0 Or Val(FMXC.txtH5.Text) > 0 Then
    FMXC.chkE.ForeColor = &HC000C0


End If
If Val(FMXC.txtH5.Text) > 0 Then
    FMXC.txtW5.Text = FMXC.txtH5.Text
End If

'fmxc.chkF.Value = mod1.HTP.Fields("chkF").Value
FMXC.txtW6.Text = Trim(Ra(28, 0))
FMXC.txtH6.Text = Trim(Ra(29, 0))
If Val(FMXC.txtW6.Text) > 0 Or Val(FMXC.txtH6.Text) > 0 Then
    FMXC.chkF.ForeColor = &HC000C0


End If
If Val(FMXC.txtH6.Text) > 0 Then
    FMXC.txtW6.Text = FMXC.txtH6.Text
End If






FMXC.txtCbze1.Text = Trim(Ra(31, 0))

FMXC.txtClcb1.Text = Trim(Ra(32, 0))

FMXC.txtRgf1.Text = Trim(Ra(33, 0))

FMXC.txtCLF1.Text = Trim(Ra(34, 0))
FMXC.txtFbje1.Text = Trim(Ra(35, 0))

FMXC.txtYf1.Text = Trim(Ra(36, 0))

FMXC.txtQt1.Text = Trim(Ra(37, 0))

FMXC.txtJlr1.Text = Trim(Ra(38, 0))
'fmxc.txtJlr2.Text = ""
FMXC.txtYj1.Text = Trim(Ra(39, 0))
'FMXC.txtYj2.Text = Trim(Ra(40, 0))
If Val(FMXC.txtYj1.Text) > 0 Then
    FMXC.optY1.Value = True
Else
    FMXC.optY2.Value = True
End If
FMXC.txtLr1.Text = Trim(Ra(41, 0))
'fmxc.txtLr2.Text = ""

FMXC.txtTcBe.Text = Trim(Ra(42, 0))
FMXC.txtTc2.Text = Trim(Ra(43, 0))
FMXC.txtTcRQ.Text = Trim(Ra(44, 0))
FMXC.lblyjFF.Caption = Trim(Ra(45, 0))


FMXC.txtF.Text = Trim(Ra(46, 0))
FMXC.txtL.Text = Trim(Ra(47, 0))


If Trim(Ra(48, 0)) = 0 Then
    FMXC.optP.Value = True
ElseIf Trim(Ra(48, 0)) = 1 Then
    FMXC.optZ.Value = True
ElseIf Trim(Ra(48, 0)) = 9 Then
    FMXC.optG.Value = True
ElseIf Trim(Ra(48, 0)) = 2 Then
    FMXC.optW.Value = True
End If

FMXC.lblMHid.Caption = Trim(Ra(49, 0))



FMXC.lblLc.Caption = Trim(Ra(50, 0))
FMXC.lblLcRen.Caption = Trim(Ra(51, 0))
FMXC.lblLcUid.Caption = Trim(Ra(52, 0))
If FMXC.lblLc.Caption = 0 Or FMXC.lblLc.Caption = 1 Then
    FMXC.lblLcRen.Caption = FMXC.txtYwy.Text
    FMXC.lblLcUid.Caption = FMXC.txtYwy.ToolTipText
End If
FMXC.lblFwid.Caption = Trim(Ra(53, 0))



If FMXC.txtHtbh.Text = "HMNEW" And FMXC.lblLcUid.Caption = mod1.DHid Then
    FMXC.cmdHT.Visible = True
Else
    FMXC.cmdHT.Visible = False
End If

'财务评定
'FMXC.txtYjf.Text = mod1.HTP.Fields("yjRQ").Value
If Trim(Ra(54, 0)) = "1999-1-1" Then FMXC.txtYjf.Text = ""
If Trim(Ra(55, 0)) = True Then
    FMXC.chkYJF.Value = 1
Else
    FMXC.chkYJF.Value = 0
End If
If Trim(Ra(56, 0)) = True Then
    FMXC.chkJTF.Value = 1
Else
    FMXC.chkJTF.Value = 0
End If
If Trim(Ra(57, 0)) Then
    FMXC.chkQKF.Value = 1
Else
    FMXC.chkQKF.Value = 0
End If
FMXC.txtYjfBz.Text = Trim(Ra(58, 0))
FMXC.FO = Val(Ra(59, 0))
'FMXC.lblFk.Caption = Trim(Ra(60, 0))
FMXC.txtZbh.Text = Trim(Ra(61, 0))

'速达金额
FMXC.txtD1.Text = Val(Ra(63, 0))
FMXC.txtD2.Text = Val(Ra(64, 0))
FMXC.txtD3.Text = Val(Ra(65, 0))
FMXC.txtD4.Text = Val(Ra(66, 0))
FMXC.txtD5.Text = Val(Ra(67, 0))
FMXC.txtD6.Text = Val(Ra(68, 0))

FMXC.comKQY.Text = Trim(Ra(69, 0)) '跨区销售
'维保
If Val(FMXC.txtH1.Text) > 0 Then
    FMXC.txtWc.Text = Trim(Rb(0, 0))
    FMXC.txtXc.Text = Trim(Rb(1, 0))
    FMXC.chkBa.Value = Trim(Rb(2, 0))
    FMXC.chkBb.Value = Trim(Rb(3, 0))
    FMXC.chkBc.Value = Trim(Rb(4, 0))
    FMXC.txtZu.Text = Trim(Rb(5, 0))
    
    uc = UBound(RC, 2)
    FMXC.MMdtgA.Clear
    FMXC.MMdtgA.Row = 0
    FMXC.MMdtgA.Col = 1
    FMXC.MMdtgA.Text = "机组品牌"
    FMXC.MMdtgA.Col = 2
    FMXC.MMdtgA.Text = "机组型号"
    FMXC.MMdtgA.Col = 3
    FMXC.MMdtgA.Text = "机组数量"
    For oo = 1 To uc + 1
        FMXC.MMdtgA.Row = oo
        For ii = 1 To 3
            FMXC.MMdtgA.Col = ii
            FMXC.MMdtgA.Text = RC(ii, oo)
        Next
    Next
    FMXC.tabGc.TabVisible(0) = True
End If
'大修
If Val(FMXC.txtH2.Text) > 0 Then
    FMXC.txtZuD.Text = RD(0, 0)
    FMXC.txtMon.Text = RD(1, 0)
    FMXC.txtDxnr.Text = RD(2, 0)

    ud = UBound(RD, 2)
    FMXC.MMdtgB.Clear
    FMXC.MMdtgB.Row = 0
    FMXC.MMdtgB.Col = 1
    FMXC.MMdtgB.Text = "机组品牌"
    FMXC.MMdtgB.Col = 2
    FMXC.MMdtgB.Text = "机组型号"
    FMXC.MMdtgB.Col = 3
    FMXC.MMdtgB.Text = "机组数量"
    For oo = 1 To ud + 1
        FMXC.MMdtgB.Row = oo
        For ii = 1 To 3
            FMXC.MMdtgB.Col = ii
            FMXC.MMdtgB.Text = RE(ii, oo)
        Next
    Next
    FMXC.tabGc.TabVisible(1) = True
End If

'If fmxc.chkC.Value = 1 Then '工程分包
If Val(FMXC.txtW3.Text) > 0 Then '工程分包
    'FMXC.txtFbnr.Text = Rf(0, 0)
    FMXC.tabGc.TabVisible(4) = True
End If

'If fmxc.chkD.Value = 1 Then '外包
If Val(FMXC.txtW4.Text) > 0 Then '
    FMXC.txtDxnr.Text = Rg(0, 0)
    FMXC.tabGc.TabVisible(5) = True
End If


FMXC.MMdtgBao.Visible = False
FMXC.MMdtgMa.Visible = False
FMXC.MMdtgCP.Visible = False
FMXC.MMdtgCPCB.Visible = False


FMXC.cmdGx.Visible = False
'FMXC.Label54.Visible = False
'FMXC.txtCj.Visible = False
'FMXC.cmdCGX.Visible = False
If Val(FMXC.txtH5.Text) > 0 Then '配件
    FMXC.MMdtgBao.FixedRows = 1
    uh = UBound(Rh, 2)
    FMXC.MMdtgBao.Row = 0: FMXC.MMdtgBao.Col = 1
    FMXC.MMdtgBao.Text = "品种"
    FMXC.MMdtgBao.Col = 2: FMXC.MMdtgBao.Text = "机组品牌"
    FMXC.MMdtgBao.Col = 3: FMXC.MMdtgBao.Text = "机组型号"
    FMXC.MMdtgBao.Col = 4: FMXC.MMdtgBao.Text = "压缩机型号"
    FMXC.MMdtgBao.Col = 5: FMXC.MMdtgBao.Text = "出厂编号"
    FMXC.MMdtgBao.Col = 6: FMXC.MMdtgBao.Text = "机组序列号"
    FMXC.MMdtgBao.Col = 7: FMXC.MMdtgBao.Text = "零件编号"
    FMXC.MMdtgBao.Col = 8: FMXC.MMdtgBao.Text = "零件名称"
    FMXC.MMdtgBao.Col = 9: FMXC.MMdtgBao.Text = "品牌产地"
    FMXC.MMdtgBao.Col = 10: FMXC.MMdtgBao.Text = "到货期"
    FMXC.MMdtgBao.Col = 11: FMXC.MMdtgBao.Text = "数量"
    FMXC.MMdtgBao.Col = 12: FMXC.MMdtgBao.Text = "单价"
    FMXC.MMdtgBao.Col = 13: FMXC.MMdtgBao.Text = "合计"
    FMXC.MMdtgBao.Col = 14: FMXC.MMdtgBao.Text = "报价有效期"
    FMXC.MMdtgBao.Col = 15: FMXC.MMdtgBao.Text = "baoid"
    FMXC.MMdtgBao.Col = 16: FMXC.MMdtgBao.Text = "Lid"
    FMXC.MMdtgBao.Col = 17: FMXC.MMdtgBao.Text = "Llid"
    FMXC.MMdtgBao.Col = 18: FMXC.MMdtgBao.Text = "bid"
    '显示配件列表
    For oo = 1 To uh + 1
        FMXC.MMdtgBao.Row = oo
        For ii = 1 To 18
            FMXC.MMdtgBao.Col = ii
            FMXC.MMdtgBao.Text = Trim(Rh(ii - 1, oo - 1))
        Next
    Next
   
    
    FMXC.MMdtgBao.MergeCol(1) = True
    FMXC.MMdtgBao.MergeCol(2) = True
    FMXC.MMdtgBao.MergeCol(10) = True
    FMXC.MMdtgBao.MergeCol(14) = True
    FMXC.MMdtgBao.MergeCells = 3

    '显示成本表
     FMXC.MMdtgMa.Row = 0: FMXC.MMdtgMa.Col = 1: FMXC.MMdtgMa.Text = "数量"
     FMXC.MMdtgMa.Col = 2: FMXC.MMdtgMa.Text = "外包单价"
     FMXC.MMdtgMa.Col = 3: FMXC.MMdtgMa.Text = "基准单价"
     FMXC.MMdtgMa.Col = 4: FMXC.MMdtgMa.Text = "外包合计"
     FMXC.MMdtgMa.Col = 5: FMXC.MMdtgMa.Text = "基准合计"

    FMXC.tabGc.TabVisible(2) = True
    FMXC.MMdtgBao.Visible = True
    FMXC.MMdtgMa.Visible = True


    FMXC.cmdGx.Visible = True
End If

'If fmxc.chkF.Value = 1 Then '产品
If Val(FMXC.txtH6.Text) > 0 Then '产品
    FMXC.MMdtgCP.FixedRows = 1
    '显示产品列表
    ui = UBound(Ri, 2)
    FMXC.MMdtgCP.Row = 0: FMXC.MMdtgCP.Col = 1
    FMXC.MMdtgCP.Text = "品种"
    FMXC.MMdtgCP.Col = 2: FMXC.MMdtgCP.Text = "机组品牌"
    FMXC.MMdtgCP.Col = 3: FMXC.MMdtgCP.Text = "机组型号"
    FMXC.MMdtgCP.Col = 4: FMXC.MMdtgCP.Text = "压缩机型号"
    FMXC.MMdtgCP.Col = 5: FMXC.MMdtgCP.Text = "出厂编号"
    FMXC.MMdtgCP.Col = 6: FMXC.MMdtgCP.Text = "机组序列号"
    FMXC.MMdtgCP.Col = 7: FMXC.MMdtgCP.Text = "零件编号"
    FMXC.MMdtgCP.Col = 8: FMXC.MMdtgCP.Text = "零件名称"
    FMXC.MMdtgCP.Col = 9: FMXC.MMdtgCP.Text = "品牌产地"
    FMXC.MMdtgCP.Col = 10: FMXC.MMdtgCP.Text = "到货期"
    FMXC.MMdtgCP.Col = 11: FMXC.MMdtgCP.Text = "数量"
    FMXC.MMdtgCP.Col = 12: FMXC.MMdtgCP.Text = "单价"
    FMXC.MMdtgCP.Col = 13: FMXC.MMdtgCP.Text = "合计"
    FMXC.MMdtgCP.Col = 14: FMXC.MMdtgCP.Text = "报价有效期"
    FMXC.MMdtgCP.Col = 15: FMXC.MMdtgCP.Text = "baoid"
    FMXC.MMdtgCP.Col = 16: FMXC.MMdtgCP.Text = "Lid"
    FMXC.MMdtgCP.Col = 17: FMXC.MMdtgCP.Text = "Llid"
    FMXC.MMdtgCP.Col = 18: FMXC.MMdtgCP.Text = "bid"
    
    '显示配件列表
    For oo = 1 To ui + 1
        FMXC.MMdtgCP.Row = oo
        For ii = 1 To 18
            FMXC.MMdtgCP.Col = ii
            FMXC.MMdtgCP.Text = Trim(Ri(ii - 1, oo - 1))
        Next
    Next
    
    FMXC.MMdtgCP.MergeCol(1) = True
    FMXC.MMdtgCP.MergeCol(2) = True
    FMXC.MMdtgCP.MergeCol(10) = True
    FMXC.MMdtgCP.MergeCol(14) = True
    FMXC.MMdtgCP.MergeCells = 3

    '显示成本表
     FMXC.MMdtgCPCB.Row = 0: FMXC.MMdtgCPCB.Col = 1: FMXC.MMdtgCPCB.Text = "数量"
     FMXC.MMdtgCPCB.Col = 2: FMXC.MMdtgCPCB.Text = "外包单价"
     FMXC.MMdtgCPCB.Col = 3: FMXC.MMdtgCPCB.Text = "基准单价"
     FMXC.MMdtgCPCB.Col = 4: FMXC.MMdtgCPCB.Text = "外包合计"
     FMXC.MMdtgCPCB.Col = 5: FMXC.MMdtgCPCB.Text = "基准合计"
    If mod1.Bm <> "商务部" Then
        FMXC.MMdtgCPCB.ColWidth(1) = 0: FMXC.MMdtgCPCB.ColWidth(3) = 0
    End If

    FMXC.tabGc.TabVisible(3) = True
    FMXC.MMdtgCP.Visible = True
    FMXC.MMdtgCPCB.Visible = True
    'FMXC.Label54.Visible = True
    'FMXC.txtCj.Visible = True
    'FMXC.cmdCGX.Visible = True
End If



'提成
    FMXC.dtgJTf.Row = 0: FMXC.dtgJTf.Col = 1: FMXC.dtgJTf.Text = "日期"
    FMXC.dtgJTf.Col = 2: FMXC.dtgJTf.Text = "金额"
    FMXC.dtgJTf.Col = 3: FMXC.dtgJTf.Text = "备注"
    FMXC.dtgJTf.Col = 4: FMXC.dtgJTf.Text = "mid"
    For oo = 1 To uj + 1
        FMXC.dtgJTf.Row = oo
        For ii = 1 To 4
            FMXC.dtgJTf.Col = ii
            FMXC.dtgJTf.Text = Trim(Rj(ii - 1, oo - 1))
        Next
    Next
    FMXC.txtJTf.Text = Rj(5, 0)


'业绩
    FMXC.dtgyjF.Row = 0: FMXC.dtgyjF.Col = 1: FMXC.dtgyjF.Text = "日期"
    FMXC.dtgyjF.Col = 2: FMXC.dtgyjF.Text = "金额"
    FMXC.dtgyjF.Col = 3: FMXC.dtgyjF.Text = "备注"
    FMXC.dtgyjF.Col = 4: FMXC.dtgyjF.Text = "mid"
    For oo = 1 To uk + 1
        FMXC.dtgyjF.Row = oo
        For ii = 1 To 4
            FMXC.dtgyjF.Col = ii
            FMXC.dtgyjF.Text = Rj(ii - 1, oo - 1)
        Next
    Next
    FMXC.txtYjf.Text = Rk(5, 0)


''''''''''tt = "SELECT rp_dd as 日期,amtn_cls as 金额,rem as 备注 FROM TF_MON where rp_id=1 and cas_no='" & FMXC.txtHtbh.Text & "' order by rp_dd"
''''''''''
''''''''''mod1.mQk.Close
''''''''''mod1.mQk.Open tt, mod1.workTx, adOpenKeyset, adLockReadOnly, adCmdText
''''''''''If IsNull(mod1.mQk.RecordCount) = True Then
''''''''''    MsgBox ("读取数据错误2.5!")
''''''''''    Exit Sub
''''''''''End If
''''''''''
''''''''''If mod1.mQk.RecordCount = 0 Then
''''''''''    Set FMXC.dtgQkf.DataSource = mod1.mQk
''''''''''    FMXC.dtgQkf.Rows = 2
''''''''''    FMXC.dtgQkf.FixedRows = 0
''''''''''    FMXC.dtgQkf.FixedRows = 1
''''''''''
''''''''''Else
''''''''''    FMXC.dtgQkf.Rows = 2
''''''''''    FMXC.dtgQkf.FixedRows = 1
''''''''''    Set FMXC.dtgQkf.DataSource = mod1.mQk
''''''''''End If
''''''''''
'''''''''''tt = "select sum(je) as je from htpingQk where hid=" & Val(FMXC.lblMHid.Caption) & " and delf=1"
''''''''''tt = "select sum(amtn_cls) as je from tf_mon where rp_id=1 and cas_no='" & FMXC.txtHtbh.Text & "'"
''''''''''Set mod1.HTP = CreateObject("adodb.recordset")
''''''''''mod1.HTP.Open tt, mod1.workTx, adOpenKeyset, adLockReadOnly, adCmdText
''''''''''If IsNull(mod1.HTP.RecordCount) = True Then
''''''''''    MsgBox ("读取数据错误2.6!")
''''''''''    Exit Sub
''''''''''End If
''''''''''FMXC.txtQkf.Text = mod1.HTP.Fields("je").Value
''''''''''FMXC.txtZe.Text = FMXC.txtQkf.Text
''''''''''FMXC.txtEd.Text = Round(Val(FMXC.txtZe.Text) / Val(FMXC.txtHtze.Text) * 100, 2)




'打开应收款表
ul = UBound(RL, 2)
FMXC.MMdtgFk.Row = 0: FMXC.MMdtgFk.Col = 1: FMXC.MMdtgFk.Text = "应付日期"
FMXC.MMdtgFk.Col = 2: FMXC.MMdtgFk.Text = "收款额度"
FMXC.MMdtgFk.Col = 3: FMXC.MMdtgFk.Text = "应付金额"
For oo = 1 To ul + 1
    FMXC.MMdtgFk.Row = oo
    For ii = 1 To 4
        FMXC.MMdtgFk.Col = ii
        FMXC.MMdtgFk.Text = Trim(RL(ii - 1, oo - 1))
        If ii = 2 Then
            FMXC.MMdtgFk.Text = Str(Val(FMXC.MMdtgFk.Text) * 100) & "%"
        End If
    Next
Next




'打开佣金表
um = UBound(RM, 2)
FMXC.MMdtgYJ.Row = 0: FMXC.MMdtgYJ.Col = 1: FMXC.MMdtgYJ.Text = "收款额度"
FMXC.MMdtgYJ.Col = 1: FMXC.MMdtgYJ.Text = "支付金额"
For oo = 1 To um + 1
    FMXC.MMdtgYJ.Row = oo
    For ii = 1 To 3
        FMXC.MMdtgYJ.Col = ii
        FMXC.MMdtgYJ.Text = Trim(RM(ii - 1, oo - 1))
    Next
Next

FMXC.txtHtbh.ToolTipText = RP(0, 0) '电子合同

'速达入帐
uq = UBound(RQ, 2) + 1
FMXC.dtgSD.Rows = uq + 20
For oo = 1 To uq
    FMXC.dtgSD.Row = oo
    For ii = 0 To 1
        FMXC.dtgSD.Col = ii
        FMXC.dtgSD.Text = RQ(ii, oo - 1)
        If ii = 1 Then
            Je = Val(FMXC.dtgSD.Text) + Je
        End If
    Next
Next
FMXC.dtgSD.Row = 0
FMXC.dtgSD.Col = 0: FMXC.dtgSD.Text = "入帐日期"
FMXC.dtgSD.Col = 1: FMXC.dtgSD.Text = "金额"

FMXC.txtZe.Text = Je
FMXC.txtEd.Text = Str(Round(Val(FMXC.txtZe.Text) / Val(FMXC.txtHtze.Text), 2) * 100)

Call modNewHT.OAn


FMXC.frmYm.Visible = False
FMXC.frmYj.Visible = False
FMXC.Visible = True





If FMXC.optZ.Value = True Then
    FMXC.cmdMod.Enabled = False
    FMXC.cmdDel.Enabled = False
    FMXC.cmdSave.Enabled = False
End If
If mod1.DName = "马晓聪" Or mod1.DName = "乔继敏" Then
    FMXC.cmdMod.Enabled = True
    FMXC.cmdDel.Enabled = False
    FMXC.cmdSave.Enabled = False
End If

If (FMXC.optZ.Value = True Or FMXC.optW.Value = True) And FMXC.txtXYwy.ToolTipText = mod1.DHid Then
    FMXC.cmdNew.Visible = True
Else
    FMXC.cmdNew.Visible = False
End If


Call FmxcFK.Qing
'kqy2,kren,kren2,kuid,kuid2,klb0,klb,klb2
FmxcFK.comQy2.Text = Trim(Ra(72, 0))
FmxcFK.comQy3.Text = Trim(Ra(73, 0))
FmxcFK.txtRen2.Text = Trim(Ra(74, 0))
FmxcFK.txtRen3.Text = Trim(Ra(75, 0))
FmxcFK.txtRen2.ToolTipText = Trim(Ra(76, 0))
FmxcFK.txtRen3.ToolTipText = Trim(Ra(77, 0))
FmxcFK.txtBL1.Text = Trim(Ra(78, 0))
FmxcFK.txtBL2.Text = Trim(Ra(79, 0))
FmxcFK.txtBL3.Text = Trim(Ra(80, 0))
    FMXC.cmdDel.Enabled = False
    FMXC.cmdSave.Enabled = False
    FMXC.Show
    FMXC.ZOrder 0
    
If FMXC.lblHtxz.Caption = "维保" Then
    FMXC.frmDate.Visible = True
End If
Exit Sub
ERCU:
MsgBox ("出错")
End
End Sub

Public Sub OAn()
Dim adoMM As Object
Dim ZT As String
Dim oo As Integer
On Error Resume Next
Set adoMM = CreateObject("adodb.recordset")
    For oo = 20 To 1 Step -1
        Unload FMXC.cmdMQm(oo)
        Unload FMXC.lblMQM(oo)
        Unload FMXC.lblMTm(oo)
    Next
    FMXC.cmdMQm(0).Caption = ""
    FMXC.lblMTm(0).Caption = ""

      ZT = "qmrzOpen(" & mod1.BTZ & ",'" & FMXC.lblMHid.Caption & "')"
      adoMM.Close

      adoMM.Open ZT, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc



      If IsNull(adoMM.RecordCount) = True Then
        MsgBox ("出错7")
        Exit Sub
      End If
      'MsgBox ("jj")
      If adoMM.RecordCount > 0 Then
         adoMM.MoveFirst
         FMXC.cmdMQm(0).Visible = True
         FMXC.lblMQM(0).Visible = True
         FMXC.lblMTm(0).Visible = True
                  FMXC.lblMQM(0).Caption = adoMM.Fields("QLabel").Value
        If adoMM.Fields("xf").Value = True Then
         FMXC.cmdMQm(0).Caption = adoMM.Fields("Qren").Value
         FMXC.lblMTm(0).Caption = adoMM.Fields("QRQ").Value
         End If
         FMXC.cmdMQm(0).Tag = adoMM.Fields("zid").Value
         adoMM.MoveNext
         For oo = 1 To adoMM.RecordCount - 1
           Load FMXC.lblMQM(oo)
           FMXC.lblMQM(oo).Caption = ""
           Load FMXC.cmdMQm(oo)
           FMXC.cmdMQm(oo).Caption = ""
           Load FMXC.lblMTm(oo)
           FMXC.lblMTm(oo).Caption = ""
           FMXC.lblMQM(oo).Caption = adoMM.Fields("QLabel").Value
            If adoMM.Fields("xf").Value = True Then
                FMXC.cmdMQm(oo).Caption = adoMM.Fields("Qren").Value
                If FMXC.cmdMQm(oo).Caption = "南京办经理" Then
                    FMXC.cmdMQm(oo).Caption = "南京办经理"
                End If
                FMXC.lblMTm(oo).Caption = adoMM.Fields("QRQ").Value
           End If
           FMXC.cmdMQm(oo).Tag = adoMM.Fields("zid").Value
           FMXC.lblMQM(oo).Visible = True
           FMXC.cmdMQm(oo).Visible = True
           FMXC.lblMTm(oo).Visible = True
           FMXC.lblMQM(oo).Left = FMXC.lblMQM(oo - 1).Left + 1000
           FMXC.cmdMQm(oo).Left = FMXC.cmdMQm(oo - 1).Left + 1000
           FMXC.lblMTm(oo).Left = FMXC.lblMTm(oo - 1).Left + 1000
           adoMM.MoveNext
        Next
     Else
        FMXC.cmdMQm(0).Visible = False
        FMXC.lblMQM(0).Visible = False
        FMXC.lblMTm(0).Visible = False
     End If

End Sub
