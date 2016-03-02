Attribute VB_Name = "modBJD"

Public Sub XJWBLcBut(Nlb As Integer)
Dim tt As String
Dim oo As Integer
On Error Resume Next
For oo = 10 To 1 Step -1
    Unload frmWBXJ.lblTm(oo)
    Unload frmWBXJ.cmdQm(oo)
    Unload frmWBXJ.lblQM(oo)
Next
    frmWBXJ.cmdQm(0).Caption = ""
    frmWBXJ.lblTm(0).Caption = ""
    tt = "lcBut(" & Nlb & ")"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    mod1.HTP.MoveFirst
    frmWBXJ.cmdQm(0).Caption = ""
    frmWBXJ.lblQM(0).Caption = mod1.HTP.Fields("LNR").Value
    frmWBXJ.lblTm(0).Caption = ""
    mod1.HTP.MoveNext '第一个数组按钮不用添加,所以,跳到下一记录
    For oo = 1 To mod1.HTP.RecordCount - 1
        Load frmWBXJ.lblQM(oo)
        Load frmWBXJ.cmdQm(oo)
        Load frmWBXJ.lblTm(oo)
        frmWBXJ.lblQM(oo).Caption = mod1.HTP.Fields("LNR").Value
        frmWBXJ.lblQM(oo).Visible = True
        frmWBXJ.lblQM(oo).Left = frmWBXJ.lblQM(oo - 1).Left + 1100
        frmWBXJ.cmdQm(oo).Caption = ""
        frmWBXJ.cmdQm(oo).Visible = True
        frmWBXJ.cmdQm(oo).Left = frmWBXJ.cmdQm(oo - 1).Left + 1100
        frmWBXJ.lblTm(oo).Caption = ""
        frmWBXJ.lblTm(oo).Visible = True
        frmWBXJ.lblTm(oo).Left = frmWBXJ.lblTm(oo - 1).Left + 1100
        mod1.HTP.MoveNext
    Next


        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "QMRZAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@NLb") = Nlb
        mod1.cmd.Parameters("@btz") = mod1.BTZ
        mod1.cmd.Parameters("@QDBH") = frmWBXJ.lblBid.Caption
        mod1.cmd.Execute
        Set cmd = Nothing
        

End Sub
Public Sub XJWBLcNew(Nlb As Integer)
Dim tt As String
Dim oo As Integer
On Error Resume Next
For oo = 10 To 1 Step -1
    Unload frmWBXJ.lblTm(oo)
    Unload frmWBXJ.cmdQm(oo)
    Unload frmWBXJ.lblQM(oo)
Next
    frmWBXJ.cmdQm(0).Caption = ""
    frmWBXJ.lblTm(0).Caption = ""
    tt = "lcBut(" & Nlb & ")"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    mod1.HTP.MoveFirst
    frmWBXJ.cmdQm(0).Caption = ""
    frmWBXJ.lblQM(0).Caption = mod1.HTP.Fields("LNR").Value
    frmWBXJ.lblTm(0).Caption = ""
    mod1.HTP.MoveNext '第一个数组按钮不用添加,所以,跳到下一记录
    For oo = 1 To mod1.HTP.RecordCount - 1
        Load frmWBXJ.lblQM(oo)
        Load frmWBXJ.cmdQm(oo)
        Load frmWBXJ.lblTm(oo)
        frmWBXJ.lblQM(oo).Caption = mod1.HTP.Fields("LNR").Value
        frmWBXJ.lblQM(oo).Visible = True
        frmWBXJ.lblQM(oo).Left = frmWBXJ.lblQM(oo - 1).Left + 1100
        frmWBXJ.cmdQm(oo).Caption = ""
        frmWBXJ.cmdQm(oo).Visible = True
        frmWBXJ.cmdQm(oo).Left = frmWBXJ.cmdQm(oo - 1).Left + 1100
        frmWBXJ.lblTm(oo).Caption = ""
        frmWBXJ.lblTm(oo).Visible = True
        frmWBXJ.lblTm(oo).Left = frmWBXJ.lblTm(oo - 1).Left + 1100
        mod1.HTP.MoveNext
    Next


'        Set mod1.cmd = createobject("adodb.command")
'        mod1.cmd.ActiveConnection = mod1.CC
'        mod1.cmd.CommandText = "QMRZAdd"
'        mod1.cmd.CommandType = adCmdStoredProc
'        mod1.cmd.Parameters("@NLb") = Nlb
'        mod1.cmd.Parameters("@btz") = mod1.BTZ
'        mod1.cmd.Parameters("@QDBH") = frmWBXJ.lblBid.Caption
'        mod1.cmd.Execute
'        Set cmd = Nothing
        

End Sub
Public Sub XJGXLcBut(Nlb As Integer)
''''''''''Dim tt As String
''''''''''Dim oo As Integer
''''''''''On Error Resume Next
''''''''''For oo = 10 To 1 Step -1
''''''''''    Unload frmGXBj.lblTm(oo)
''''''''''    Unload frmGXBj.cmdQm(oo)
''''''''''    Unload frmGXBj.lblQM(oo)
''''''''''Next
''''''''''    frmGXBj.cmdQm(0).Caption = ""
''''''''''    frmGXBj.lblTm(0).Caption = ""
''''''''''    tt = "lcBut(" & Nlb & ")"
''''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
''''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
''''''''''    mod1.HTP.MoveFirst
''''''''''    frmGXBj.cmdQm(0).Caption = ""
''''''''''    frmGXBj.lblQM(0).Caption = mod1.HTP.Fields("LNR").Value
''''''''''    frmGXBj.lblTm(0).Caption = ""
''''''''''    mod1.HTP.MoveNext '第一个数组按钮不用添加,所以,跳到下一记录
''''''''''    For oo = 1 To mod1.HTP.RecordCount - 1
''''''''''        Load frmGXBj.lblQM(oo)
''''''''''        Load frmGXBj.cmdQm(oo)
''''''''''        Load frmGXBj.lblTm(oo)
''''''''''        frmGXBj.lblQM(oo).Caption = mod1.HTP.Fields("LNR").Value
''''''''''        frmGXBj.lblQM(oo).Visible = True
''''''''''        frmGXBj.lblQM(oo).Left = frmGXBj.lblQM(oo - 1).Left + 1100
''''''''''        frmGXBj.cmdQm(oo).Caption = ""
''''''''''        frmGXBj.cmdQm(oo).Visible = True
''''''''''        frmGXBj.cmdQm(oo).Left = frmGXBj.cmdQm(oo - 1).Left + 1100
''''''''''        frmGXBj.lblTm(oo).Caption = ""
''''''''''        frmGXBj.lblTm(oo).Visible = True
''''''''''        frmGXBj.lblTm(oo).Left = frmGXBj.lblTm(oo - 1).Left + 1100
''''''''''        mod1.HTP.MoveNext
''''''''''    Next
''''''''''
''''''''''
''''''''''        Set mod1.cmd = createobject("adodb.command")
''''''''''        mod1.cmd.ActiveConnection = mod1.cc
''''''''''        mod1.cmd.CommandText = "QMRZAdd"
''''''''''        mod1.cmd.CommandType = adCmdStoredProc
''''''''''        mod1.cmd.Parameters("@NLb") = Nlb
''''''''''        mod1.cmd.Parameters("@btz") = mod1.BTZ
''''''''''        mod1.cmd.Parameters("@QDBH") = frmGXBj.lblBid.Caption
''''''''''        mod1.cmd.Execute
''''''''''        Set cmd = Nothing
        

End Sub
Public Sub XJGXLcNew(Nlb As Integer)
''''''Dim tt As String
''''''Dim oo As Integer
''''''On Error Resume Next
''''''For oo = 10 To 1 Step -1
''''''    Unload frmGXBj.lblTm(oo)
''''''    Unload frmGXBj.cmdQm(oo)
''''''    Unload frmGXBj.lblQM(oo)
''''''Next
''''''    frmGXBj.cmdQm(0).Caption = ""
''''''    frmGXBj.lblTm(0).Caption = ""
''''''    tt = "lcBut(" & Nlb & ")"
''''''    Set mod1.HTP = CreateObject("adodb.recordset")
''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
''''''    If IsNull(mod1.HTP.RecordCount) = True Then
''''''        MsgBox ("读取数据有误,请在关闭后,再试一次!")
''''''        End
''''''    End If
''''''    mod1.HTP.MoveFirst
''''''    frmGXBj.cmdQm(0).Caption = ""
''''''    frmGXBj.lblQM(0).Caption = mod1.HTP.Fields("LNR").Value
''''''    frmGXBj.lblTm(0).Caption = ""
''''''    mod1.HTP.MoveNext '第一个数组按钮不用添加,所以,跳到下一记录
''''''    For oo = 1 To mod1.HTP.RecordCount - 1
''''''        Load frmGXBj.lblQM(oo)
''''''        Load frmGXBj.cmdQm(oo)
''''''        Load frmGXBj.lblTm(oo)
''''''        frmGXBj.lblQM(oo).Caption = mod1.HTP.Fields("LNR").Value
''''''        frmGXBj.lblQM(oo).Visible = True
''''''        frmGXBj.lblQM(oo).Left = frmGXBj.lblQM(oo - 1).Left + 1100
''''''        frmGXBj.cmdQm(oo).Caption = ""
''''''        frmGXBj.cmdQm(oo).Visible = True
''''''        frmGXBj.cmdQm(oo).Left = frmGXBj.cmdQm(oo - 1).Left + 1100
''''''        frmGXBj.lblTm(oo).Caption = ""
''''''        frmGXBj.lblTm(oo).Visible = True
''''''        frmGXBj.lblTm(oo).Left = frmGXBj.lblTm(oo - 1).Left + 1100
''''''        mod1.HTP.MoveNext
''''''    Next


'        Set mod1.cmd = createobject("adodb.command")
'        mod1.cmd.ActiveConnection = mod1.CC
'        mod1.cmd.CommandText = "QMRZAdd"
'        mod1.cmd.CommandType = adCmdStoredProc
'        mod1.cmd.Parameters("@NLb") = Nlb
'        mod1.cmd.Parameters("@btz") = mod1.BTZ
'        mod1.cmd.Parameters("@QDBH") = frmGXBj.lblBid.Caption
'        mod1.cmd.Execute
'        Set cmd = Nothing
        

End Sub
Public Sub BjWBLcBut(Nlb As Integer)
Dim tt As String
Dim oo As Integer
On Error Resume Next
For oo = 10 To 1 Step -1
    Unload frmWbxjB.lblTm(oo)
    Unload frmWbxjB.cmdQm(oo)
    Unload frmWbxjB.lblQM(oo)
Next
    frmWbxjB.cmdQm(0).Caption = ""
    frmWbxjB.lblTm(0).Caption = ""
    frmWbxjB.cmdQm(0).Visible = True
    frmWbxjB.lblQM(0).Visible = True
    frmWbxjB.lblTm(0).Visible = True
    tt = "lcBut(" & Nlb & ")"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    mod1.HTP.MoveFirst
    frmWbxjB.cmdQm(0).Caption = ""
    frmWbxjB.lblQM(0).Caption = mod1.HTP.Fields("LNR").Value
    frmWbxjB.lblTm(0).Caption = ""
    mod1.HTP.MoveNext '第一个数组按钮不用添加,所以,跳到下一记录
    For oo = 1 To mod1.HTP.RecordCount - 1
        Load frmWbxjB.lblQM(oo)
        Load frmWbxjB.cmdQm(oo)
        Load frmWbxjB.lblTm(oo)
        frmWbxjB.lblQM(oo).Caption = mod1.HTP.Fields("LNR").Value
        frmWbxjB.lblQM(oo).Visible = True
        frmWbxjB.lblQM(oo).Left = frmWbxjB.lblQM(oo - 1).Left + 1100
        frmWbxjB.cmdQm(oo).Caption = ""
        frmWbxjB.cmdQm(oo).Visible = True
        frmWbxjB.cmdQm(oo).Left = frmWbxjB.cmdQm(oo - 1).Left + 1100
        frmWbxjB.lblTm(oo).Caption = ""
        frmWbxjB.lblTm(oo).Visible = True
        frmWbxjB.lblTm(oo).Left = frmWbxjB.lblTm(oo - 1).Left + 1100
        mod1.HTP.MoveNext
    Next


        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "QMRZAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@NLb") = Nlb
        mod1.cmd.Parameters("@btz") = mod1.BTZ
        mod1.cmd.Parameters("@QDBH") = frmWbxjB.lblBaoId.Caption
        mod1.cmd.Execute
        Set cmd = Nothing
        

End Sub
Public Sub BjGXLcBut(Nlb As Integer)
Dim tt As String
Dim oo As Integer
On Error Resume Next
For oo = 10 To 1 Step -1
    Unload frmGxbjB.lblTm(oo)
    Unload frmGxbjB.cmdQm(oo)
    Unload frmGxbjB.lblQM(oo)
Next
    frmGxbjB.cmdQm(0).Caption = ""
    frmGxbjB.lblTm(0).Caption = ""
    frmGxbjB.cmdQm(0).Visible = True
    frmGxbjB.lblQM(0).Visible = True
    frmGxbjB.lblTm(0).Visible = True
    tt = "lcBut(" & Nlb & ")"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    mod1.HTP.MoveFirst
    frmGxbjB.cmdQm(0).Caption = ""
    frmGxbjB.lblQM(0).Caption = mod1.HTP.Fields("LNR").Value
    frmGxbjB.lblTm(0).Caption = ""
    mod1.HTP.MoveNext '第一个数组按钮不用添加,所以,跳到下一记录
    For oo = 1 To mod1.HTP.RecordCount - 1
        Load frmGxbjB.lblQM(oo)
        Load frmGxbjB.cmdQm(oo)
        Load frmGxbjB.lblTm(oo)
        frmGxbjB.lblQM(oo).Caption = mod1.HTP.Fields("LNR").Value
        frmGxbjB.lblQM(oo).Visible = True
        frmGxbjB.lblQM(oo).Left = frmGxbjB.lblQM(oo - 1).Left + 1100
        frmGxbjB.cmdQm(oo).Caption = ""
        frmGxbjB.cmdQm(oo).Visible = True
        frmGxbjB.cmdQm(oo).Left = frmGxbjB.cmdQm(oo - 1).Left + 1100
        frmGxbjB.lblTm(oo).Caption = ""
        frmGxbjB.lblTm(oo).Visible = True
        frmGxbjB.lblTm(oo).Left = frmGxbjB.lblTm(oo - 1).Left + 1100
        mod1.HTP.MoveNext
    Next


        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "QMRZAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@NLb") = Nlb
        mod1.cmd.Parameters("@btz") = mod1.BTZ
        mod1.cmd.Parameters("@QDBH") = frmGxbjB.lblBaoId.Caption
        mod1.cmd.Execute
        Set cmd = Nothing
        

End Sub
Public Sub BJDWBQing() '维保询价单清空
Dim tt As String
On Error Resume Next
'tt = "select jzpb as 机组品牌,jzxh as 机组型号,xt as 系统类别,wnr as 工作内容,gt as 工时,dgt as 附加工时,dl as 附加量,dw as 单位,bf as 设备费,bz as 备注,xg as 合计,wbx,fjl,bid,wid,lid from bjxt"
'frmWBXJ.adoWb.Close
'frmWBXJ.adoWb.Open tt, mod1.workBD, adOpenKeyset, adLockReadOnly, adCmdText

Set frmWBXJ.dtgWb.DataSource = Nothing
Set frmWBXJ.dtgLj.DataSource = Nothing
frmWBXJ.lblZl.Caption = ""
frmWBXJ.dtgLj.Visible = False
frmWBXJ.dtgWb.Visible = True
frmWBXJ.comXmmc.Tag = ""
frmWBXJ.comXmmc.Text = ""
frmWBXJ.lblBid.Caption = ""
frmWBXJ.lblBh.Caption = ""
frmWBXJ.comZu.Text = ""
frmWBXJ.txtZu.Text = ""
frmWBXJ.comPb.Text = ""
frmWBXJ.comXh.Text = ""
frmWBXJ.txtSL.Text = ""
frmWBXJ.lblOid.Caption = ""
frmWBXJ.txtZT.Text = ""
frmWBXJ.txtHg.Text = ""
frmWBXJ.txtYhg.Text = ""
frmWBXJ.chkBa.Value = 0
frmWBXJ.chkBb.Value = 0
frmWBXJ.chkBc.Value = 0
frmWBXJ.lblYwy.Caption = ""
frmWBXJ.lblUid.Caption = ""
frmWBXJ.cmdLeft.Enabled = False
frmWBXJ.cmdRight.Enabled = False
frmWBXJ.txtWc.Text = ""
frmWBXJ.txtXc.Text = ""
frmWBXJ.lblCgid.Caption = ""
frmWBXJ.lblLc.Caption = ""
frmWBXJ.lblLcRen.Caption = ""
frmWBXJ.lblLcUid.Caption = ""
frmWBXJ.lblFwid.Caption = ""
frmWBXJ.lblNlb.Caption = ""
frmWBXJ.lblLcou.Caption = "" '
frmWBXJ.lblPwf.Caption = ""
frmWBXJ.lblBaoId.Caption = ""
frmWBXJ.txtMon.Text = ""
frmWBXJ.txtDxnr.Text = ""
frmWBXJ.lblQy.Caption = ""
frmWBXJ.lblBM.Caption = ""
frmWBXJ.txtClf.Text = ""
frmWBXJ.txtF1.Text = ""
frmWBXJ.txtF2.Text = ""
frmWBXJ.txtF3.Text = ""
frmWBXJ.txtF4.Text = ""
frmWBXJ.txtBz.Text = ""
'frmWBXJ.cmdCong.Visible = False
frmWBXJ.txtFbje.Text = ""
frmWBXJ.txtFbnr.Text = ""
frmWBXJ.lblHtbh.Caption = ""
frmWBXJ.JZ = 0
frmWBXJ.txt1.Text = ""
frmWBXJ.txt2.Text = ""
frmWBXJ.lblHLC.Caption = ""

frmWBXJ.frmM1.Visible = False
frmWBXJ.frmM2.Visible = False
frmWBXJ.frmM3.Visible = False
frmWBXJ.frmM5.Visible = False
frmWBXJ.cmdDel.Enabled = False
End Sub
Public Sub BJDGXQing() '询价单清空
Dim tt As String
On Error Resume Next
'tt = "select jzpb as 机组品牌,jzxh as 机组型号,xt as 系统类别,wnr as 工作内容,gt as 工时,dgt as 附加工时,dl as 附加量,dw as 单位,bf as 设备费,bz as 备注,xg as 合计,wbx,fjl,bid,wid,lid from bjxt"
'frmGXBj.adoWb.Close
'frmGXBj.adoWb.Open tt, mod1.workBD, adOpenKeyset, adLockReadOnly, adCmdText
frmGXBj.txtZBQ.Text = ""
Set frmGXBj.dtgMa.DataSource = Nothing
frmGXBj.dtgMa.Clear: frmGXBj.dtgMa.Cols = 2: frmGXBj.dtgMa.FixedCols = 1
frmGXBj.comXmmc.Text = ""
frmGXBj.comXmmc.Tag = ""
frmGXBj.comJzPb1.Text = ""
frmGXBj.txtJzxh.Text = ""
frmGXBj.comJzpb.Text = ""
frmGXBj.comJzXh.Text = ""
frmGXBj.txtYxh.Text = ""
frmGXBj.txtCbh.Text = ""
frmGXBj.txtCd.Text = ""
frmGXBj.txtLjbh.Text = ""
frmGXBj.txtLjmc.Text = ""
frmGXBj.txtXlh.Text = ""
frmGXBj.txtSL.Text = ""
frmGXBj.lblBh.Caption = ""
frmGXBj.txtDj.Text = ""
frmGXBj.txtBrq.Text = ""
frmGXBj.txtMj.Text = ""
frmGXBj.txtHtbh.Text = ""
frmGXBj.lblBid.Caption = ""
frmGXBj.lblOid.Caption = ""
frmGXBj.lblHLC.Caption = ""
frmGXBj.lblLc.Caption = ""
frmGXBj.lblLcRen.Caption = ""
frmGXBj.lblLcUid.Caption = ""
frmGXBj.lblFwid.Caption = ""
frmGXBj.lblNlb.Caption = ""
frmGXBj.lblLcou.Caption = "" '
frmGXBj.lblBaoId.Caption = ""
frmGXBj.lblPwf.Caption = ""
frmGXBj.lblWhg.Caption = ""
frmGXBj.txtHg.Text = ""
frmGXBj.txtYhg.Text = ""
frmGXBj.lblWbid.Caption = ""
frmGXBj.txtBz.Text = ""
frmGXBj.lblYwy.Caption = ""
frmGXBj.lblUid.Caption = ""
frmGXBj.cmdLeft.Enabled = False
frmGXBj.cmdRight.Enabled = False
frmGXBj.txtYhg.Locked = True
'frmGXBj.lblPz.ForeColor = &H80000012
'frmGXBj.comLx.ForeColor = &H80000012
frmGXBj.lblZl.ForeColor = &H80000012
frmGXBj.lblzlZ.ForeColor = &H80000012
frmGXBj.frmZ.Visible = False
frmGXBj.frmCT.Visible = False
frmGXBj.cmdCT.Caption = ""
frmGXBj.lblCT.Caption = ""

frmGXBj.txtYf.Text = ""
frmGXBj.txtADR.Text = ""
frmGXBj.lblZ.Visible = False
frmGXBj.lblHtbh.Caption = ""
frmGXBj.lblZT.Visible = False
frmGXBj.lblZT.Caption = ""
frmGXBj.lblCfwid.Caption = ""
frmGXBj.frmSd.Visible = False
frmGXBj.txtGyid.Text = ""
frmGXBj.txtGymc.Text = ""
frmGXBj.txtGyman.Text = ""
frmGXBj.txtGyAdr.Text = ""
frmGXBj.txtGYPho.Text = ""
frmGXBj.txtGybz.Text = ""
frmGXBj.JZ = 0
frmGXBj.txtJdj.Text = ""
If mod1.Bm = "零件事业部" Or mod1.DName = "徐瑛" Or mod1.Bm = "北京配送中心" Then
    frmGXBj.frmJ.Visible = False
Else
    frmGXBj.frmJ.Visible = True
End If
frmGXBj.cmdD.Enabled = False
frmGXBj.txtFj.Text = ""
frmGXBj.txtFj.Locked = True
frmGXBj.lblSDJE.Caption = 0
'Call frmGXBj.dtgMaFF
Call frmGXBj.dtgPFF
End Sub


Public Sub BaoJDWBQing() '维保报价单清空
Dim tt As String
On Error Resume Next
'tt = "select jzpb as 机组品牌,jzxh as 机组型号,xt as 系统类别,wnr as 工作内容,gt as 工时,dgt as 附加工时,dl as 附加量,dw as 单位,bf as 设备费,bz as 备注,xg as 合计,wbx,fjl,bid,wid,lid from bjxt"
'frmwbxjb.adoWb.Close
'frmwbxjb.adoWb.Open tt, mod1.workBD, adOpenKeyset, adLockReadOnly, adCmdText

Set frmWbxjB.dtgWb.DataSource = Nothing
Set frmWbxjB.dtgLj.DataSource = Nothing
Set frmWbxjB.dtgYJ.DataSource = Nothing
Set frmWbxjB.dtgGD.DataSource = Nothing
frmWbxjB.optGDA.Value = False
frmWbxjB.optGDB.Value = False
frmWbxjB.optGDC.Value = False

frmWbxjB.comXmmc.Tag = ""
frmWbxjB.comXmmc.Text = ""
frmWbxjB.comKhmc.Text = ""
frmWbxjB.comKhmc.ToolTipText = ""
frmWbxjB.lblBid.Caption = ""
frmWbxjB.lblBh.Caption = ""
frmWbxjB.comZu.Text = ""
frmWbxjB.txtZu.Text = ""
frmWbxjB.comPb.Text = ""
frmWbxjB.comXh.Text = ""
frmWbxjB.txtSL.Text = ""
frmWbxjB.lblOid.Caption = ""
frmWbxjB.txtZT.Text = ""
frmWbxjB.txtHg.Text = ""
frmWbxjB.txtYhg.Text = ""
frmWbxjB.chkBa.Value = 0
frmWbxjB.chkBb.Value = 0
frmWbxjB.chkBc.Value = 0
frmWbxjB.lblYwy.Caption = ""
frmWbxjB.lblUid.Caption = ""
frmWbxjB.cmdLeft.Enabled = False
frmWbxjB.cmdRight.Enabled = False
frmWbxjB.txtFbje.Text = ""
frmWbxjB.txtTl.Text = ""
frmWbxjB.txtFbnr.Text = ""
frmWbxjB.lblLc.Caption = ""
frmWbxjB.lblLcRen.Caption = ""
frmWbxjB.lblLcUid.Caption = ""
frmWbxjB.lblFwid.Caption = ""
frmWbxjB.lblNlb.Caption = ""
frmWbxjB.lblLcou.Caption = "" '
frmWbxjB.lblBaoId.Caption = ""
frmWbxjB.txtRgf.Text = ""
frmWbxjB.txtClf.Text = ""
frmWbxjB.txtClcb.Text = ""
frmWbxjB.txtYJ.Text = ""
frmWbxjB.txtMon.Text = ""
frmWbxjB.txtWc.Text = ""
frmWbxjB.txtXc.Text = ""
frmWbxjB.txtF.Text = ""
frmWbxjB.txtL.Text = ""
frmWbxjB.txtTcBe.Text = ""
frmWbxjB.txtXm1.Text = ""
frmWbxjB.txtXm2.Text = ""
frmWbxjB.txtYf.Text = ""
frmWbxjB.txtYJ.Text = ""
frmWbxjB.txtTcBe.Text = ""
frmWbxjB.txtF.Text = ""
frmWbxjB.txtL.Text = ""
frmWbxjB.optLa.Value = False
frmWbxjB.optLb.Value = False
frmWbxjB.optLc.Value = False
frmWbxjB.txtBz.Text = ""
End Sub
Public Sub BaoJDGXQing() '购销报价单清空
Dim tt As String
On Error Resume Next
'tt = "select jzpb as 机组品牌,jzxh as 机组型号,xt as 系统类别,wnr as 工作内容,gt as 工时,dgt as 附加工时,dl as 附加量,dw as 单位,bf as 设备费,bz as 备注,xg as 合计,wbx,fjl,bid,wid,lid from bjxt"
'frmgxbjb.adoWb.Close
'frmgxbjb.adoWb.Open tt, mod1.workBD, adOpenKeyset, adLockReadOnly, adCmdText

Set frmGxbjB.dtgMa.DataSource = Nothing
Set frmGxbjB.dtgBao.DataSource = Nothing
Set frmGxbjB.dtgYJ.DataSource = Nothing
frmGxbjB.comXmmc.Text = ""
frmGxbjB.comXmmc.Tag = ""
frmGxbjB.comKhmc.Text = ""
frmGxbjB.comKhmc.ToolTipText = ""
frmGxbjB.txtDj.Text = ""
frmGxbjB.txtSL.Text = ""

Set frmGxbjB.dtgGD.DataSource = Nothing
frmGxbjB.optGDA.Value = False
frmGxbjB.optGDB.Value = False
frmGxbjB.optGDC.Value = False
frmGxbjB.txtGDNR.Text = ""
frmGxbjB.txtQdj.Text = ""
frmGxbjB.txtRl.Text = ""


frmGxbjB.lblBid.Caption = ""
frmGxbjB.lblOid.Caption = ""
frmGxbjB.txtFbje.Text = ""
frmGxbjB.lblLc.Caption = ""
frmGxbjB.lblLcRen.Caption = ""
frmGxbjB.lblLcUid.Caption = ""
frmGxbjB.lblFwid.Caption = ""
frmGxbjB.lblNlb.Caption = ""
frmGxbjB.lblLcou.Caption = "" '
frmGxbjB.lblBaoId.Caption = ""
frmGxbjB.txtTcBe.Text = ""

frmGxbjB.txtHg.Text = ""
frmGxbjB.txtYhg.Text = ""


frmGxbjB.lblYwy.Caption = ""
frmGxbjB.lblUid.Caption = ""
frmGxbjB.cmdLeft.Enabled = False
frmGxbjB.cmdRight.Enabled = False

frmGxbjB.txtXm1.Text = ""
frmGxbjB.txtXm2.Text = ""
frmGxbjB.txtClcb.Text = ""
frmGxbjB.txtYJ.Text = ""
frmGxbjB.txtYf.Text = ""
frmGxbjB.txtCb.Text = ""
frmGxbjB.lblHtbh.Caption = ""
frmGxbjB.txtYJ.Text = ""
frmGxbjB.txtTcBe.Text = ""
frmGxbjB.optLa.Value = False
frmGxbjB.optLb.Value = False
frmGxbjB.optLc.Value = False
frmGxbjB.txtBz.Text = ""
End Sub
Public Sub BJDBound(Bid As Long, ZL As String)   '
Dim tt As String
Dim LX As Boolean
Dim oo As Integer
On Error Resume Next
mod1.BTZ = 36
'tt="select newf from htping where hid=" &
tt = "select top 1 * from XunJiaD where bid=" & Bid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.HTP.RecordCount) = True Then
    MsgBox ("连接超时,将退出!")
    End
End If

'''If mod1.CFT = 0 And mod1.DName <> "徐瑛" Then
'''    frmGXBj.cmdFk.Visible = False
'''Else
'''    frmGXBj.cmdFk.Visible = True
'''End If

If ZL <> "购销" And ZL <> "配件" And ZL <> "产品" And ZL <> "配件询价单" And ZL <> "零配件" And ZL <> "三菱" And ZL <> "松下" And ZL <> "勤达富" And ZL <> "德图" And ZL <> "非代理产品" Then '维保或大修,分包

    If mod1.Bm = "零件事业部" Then '如果是采购,则只能看到相应的购销单
        LX = False
        tt = "select * from xunJiaD where bid=" & mod1.HTP.Fields("cgid").Value
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        frmGXBj.lblZl.Caption = mod1.HTP.Fields("Zl").Value
        frmGXBj.comXmmc.Text = mod1.HTP.Fields("xmmc").Value
        frmGXBj.comXmmc.Tag = mod1.HTP.Fields("xid").Value
        frmGXBj.lblBid.Caption = mod1.HTP.Fields("bid").Value
        frmGXBj.lblOid.Caption = mod1.HTP.Fields("oid").Value
        frmGXBj.lblLc.Caption = mod1.HTP.Fields("lc").Value
        frmGXBj.lblLcRen.Caption = mod1.HTP.Fields("lcren").Value
        frmGXBj.lblLcUid.Caption = mod1.HTP.Fields("lcuid").Value
        frmGXBj.lblFwid.Caption = mod1.HTP.Fields("fwid").Value
        frmGXBj.lblNlb.Caption = mod1.HTP.Fields("nlb").Value
        frmGXBj.lblLcou.Caption = mod1.HTP.Fields("lcou").Value
        frmGXBj.lblBaoId.Caption = mod1.HTP.Fields("baoid").Value
        frmGXBj.lblBh.Caption = mod1.HTP.Fields("bianhao").Value
        frmGXBj.lblPwf.Caption = mod1.HTP.Fields("pwf").Value
        frmGXBj.txtHg.Text = mod1.HTP.Fields("hg").Value
        frmGXBj.LBLhG.Caption = mod1.HTP.Fields("hg").Value
        frmGXBj.txtYhg.Text = mod1.HTP.Fields("yhg").Value
        frmGXBj.LBLyHG.Caption = mod1.HTP.Fields("yhg").Value
        frmGXBj.lblWhg.Caption = mod1.HTP.Fields("whg").Value
        frmGXBj.lblYwy.Caption = mod1.HTP.Fields("ywy").Value
        frmGXBj.lblUid.Caption = mod1.HTP.Fields("uid").Value
        frmGXBj.lblWbid.Caption = mod1.HTP.Fields("wbid").Value
        frmGXBj.txtBz.Text = mod1.HTP.Fields("bz").Value
        frmGXBj.ZF = mod1.HTP.Fields("zf").Value
        frmGXBj.txtHtbh.Text = mod1.HTP.Fields("htbh").Value
        frmGXBj.txtYf.Text = mod1.HTP.Fields("yf").Value
        frmGXBj.txtADR.Text = mod1.HTP.Fields("yfadr").Value
    If mod1.HTP.Fields("chf").Value = True And frmGXBj.lblLc.Caption > 2 Then
        frmGXBj.lblZ.Visible = True
        frmGXBj.lblZT.Visible = True
        frmGXBj.lblZT.Caption = mod1.HTP.Fields("chdate").Value

    End If
        frmGXBj.lblCfwid.Caption = mod1.HTP.Fields("cfwid").Value
        If frmGXBj.lblZl.Caption = "大修" Then
            frmGXBj.cmdCT.Caption = mod1.HTP.Fields("CC").Value
            frmGXBj.lblCT.Caption = mod1.HTP.Fields("ctime").Value
            frmGXBj.frmCT.Visible = True
        End If
        tt = "select qy,bm from RenYuan where userName='" & frmGXBj.lblYwy.Caption & "' and userid='" & frmGXBj.lblUid.Caption & "'"
        mod1.HTT.Close
        mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        frmGXBj.lblBM.Caption = mod1.HTT.Fields("bm").Value
        frmGXBj.lblQy.Caption = mod1.HTT.Fields("qy").Value
        
        
        tt = "select * from xunJIamxView where bid=" & Val(frmGXBj.lblBid.Caption)
        frmGXBj.adoGx.Close
        frmGXBj.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmGXBj.dtgMa.DataSource = frmGXBj.adoGx
        If frmGXBj.adoGx.RecordCount > 0 Then
            frmGXBj.dtgMa.FixedRows = 0
            frmGXBj.dtgMa.MergeCol(1) = True
            frmGXBj.dtgMa.MergeCol(2) = True
            frmGXBj.dtgMa.MergeCol(10) = True
            frmGXBj.dtgMa.MergeCol(14) = True
            frmGXBj.dtgMa.MergeCells = 3
            frmGXBj.dtgMa.FixedRows = 1
        End If
        frmGXBj.cmdSave.Enabled = False
        frmGXBj.cmdMod.Enabled = True

        frmGXBj.cmdBjd.Visible = False
        frmGXBj.cmdMod.Enabled = True
        frmGXBj.cmdSave.Enabled = False
        Call modBJD.OpenXJAN(LX)
        
    Else
        LX = True
        frmWBXJ.JZ = mod1.HTP.Fields("jz").Value
        frmWBXJ.tabGc.TabCaption(2) = mod1.HTP.Fields("Zl").Value
        frmWBXJ.cmdCG.Visible = True
        frmWBXJ.lblZl.Caption = mod1.HTP.Fields("Zl").Value
        frmWBXJ.comXmmc.Text = mod1.HTP.Fields("xmmc").Value
        frmWBXJ.comXmmc.Tag = mod1.HTP.Fields("xid").Value
        frmWBXJ.lblBid.Caption = mod1.HTP.Fields("bid").Value
        'bid = frmWBXJ.lblBid.Caption
        frmWBXJ.lblBh.Caption = mod1.HTP.Fields("bianhao").Value
        frmWBXJ.comZu.Text = mod1.HTP.Fields("zh").Value
        frmWBXJ.txtZu.Text = mod1.HTP.Fields("zName").Value
        frmWBXJ.comPb.Text = mod1.HTP.Fields("jzpb").Value
        frmWBXJ.comXh.Text = mod1.HTP.Fields("jzxh").Value
        frmWBXJ.txtSL.Text = mod1.HTP.Fields("sL").Value
        frmWBXJ.lblOid.Caption = mod1.HTP.Fields("oid").Value
        frmWBXJ.txtZT.Text = mod1.HTP.Fields("ZTime").Value
        frmWBXJ.txtHg.Text = mod1.HTP.Fields("HG").Value
        frmWBXJ.txtYhg.Text = mod1.HTP.Fields("yhg").Value
        frmWBXJ.txt1.Text = mod1.HTP.Fields("HG").Value
        frmWBXJ.txt2.Text = mod1.HTP.Fields("jhg").Value
        frmWBXJ.txtClf.Text = mod1.HTP.Fields("clf").Value
        frmWBXJ.chkBa.Value = Abs(CInt(mod1.HTP.Fields("ta").Value))
        frmWBXJ.chkBb.Value = Abs(CInt(mod1.HTP.Fields("tb").Value))
        frmWBXJ.chkBc.Value = Abs(CInt(mod1.HTP.Fields("tc").Value))
        frmWBXJ.lblYwy.Caption = mod1.HTP.Fields("ywy").Value
        frmWBXJ.lblUid.Caption = mod1.HTP.Fields("uid").Value
        frmWBXJ.txtF1.Text = mod1.HTP.Fields("f1").Value
        frmWBXJ.txtF2.Text = mod1.HTP.Fields("f2").Value
        frmWBXJ.txtF3.Text = mod1.HTP.Fields("f3").Value
        frmWBXJ.txtF4.Text = mod1.HTP.Fields("f4").Value
        tt = "select qy,bm from RenYuan where userName='" & frmWBXJ.lblYwy.Caption & "' and userid='" & frmWBXJ.lblUid.Caption & "'"
        mod1.HTT.Close
        mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        frmWBXJ.lblBM.Caption = mod1.HTT.Fields("bm").Value
        frmWBXJ.lblQy.Caption = mod1.HTT.Fields("qy").Value
        frmWBXJ.lblBaoId.Caption = mod1.HTP.Fields("baoid").Value
        frmWBXJ.txtWc.Text = mod1.HTP.Fields("wc").Value
        frmWBXJ.txtXc.Text = mod1.HTP.Fields("Xc").Value
        frmWBXJ.txtMon.Text = mod1.HTP.Fields("mon").Value
        frmWBXJ.txtDxnr.Text = mod1.HTP.Fields("dxnr").Value
        frmWBXJ.lblCgid.Caption = mod1.HTP.Fields("cgid").Value
        frmWBXJ.lblPwf.Caption = mod1.HTP.Fields("pwf").Value
        frmWBXJ.lblLc.Caption = mod1.HTP.Fields("Lc").Value
        frmWBXJ.lblLcRen.Caption = mod1.HTP.Fields("LcRen").Value
        frmWBXJ.lblLcUid.Caption = mod1.HTP.Fields("LcUid").Value
        frmWBXJ.lblFwid.Caption = mod1.HTP.Fields("Fwid").Value
        frmWBXJ.lblNlb.Caption = mod1.HTP.Fields("Nlb").Value
        frmWBXJ.lblLcou.Caption = mod1.HTP.Fields("Lcou").Value
        frmWBXJ.ZF = mod1.HTP.Fields("zf").Value
        frmWBXJ.txtBz.Text = mod1.HTP.Fields("bz").Value
        frmWBXJ.txtFbje.Text = mod1.HTP.Fields("fbje").Value
        frmWBXJ.txtFbnr.Text = mod1.HTP.Fields("fbnr").Value
        frmWBXJ.lblHtbh.Caption = mod1.HTP.Fields("htbh").Value

        If mod1.HTP.Fields("pwf").Value = True Then         '报价单按钮是否显示
            'frmWBXJ.cmdBjd.Visible = True
        Else
            frmWBXJ.cmdBjd.Visible = False
        End If
        If Val(frmWBXJ.lblBid.Caption) >= 6794 Or Val(frmWBXJ.lblBid.Caption) = 6611 Then
            frmWBXJ.frmOld.Visible = False
            frmWBXJ.frmN.Visible = True
            
            If mod1.Bm = "配送中心" Then
                frmWBXJ.lbl1.Visible = True: frmWBXJ.txt1.Visible = True
                frmWBXJ.lbl2.Visible = True: frmWBXJ.txt2.Visible = True
            Else
                frmWBXJ.lbl1.Visible = False: frmWBXJ.txt1.Visible = False
                frmWBXJ.lbl2.Visible = True: frmWBXJ.txt2.Visible = True
                If mod1.Bm = "商务部" And (frmWBXJ.lblZl.Caption = "水处理" Or frmWBXJ.lblZl.Caption = "工程分包") Then
                    frmWBXJ.lbl1.Visible = True: frmWBXJ.txt1.Visible = True
                End If
            End If
        Else
            frmWBXJ.frmOld.Visible = True
            frmWBXJ.frmN.Visible = False
        End If
        If mod1.HTP.Fields("zl").Value = "维保" Then
        
        
        
            If frmWBXJ.comPb.Text <> "" Then '如果为旧版本,则显示年保和例检条款.
                '年保表
                tt = "select * from xunJIaWbView where wbx='年保' and bid=" & Val(frmWBXJ.lblBid.Caption)
            Else
                tt = "select * from xunJIaWbView where wbx='年保' and bid=0"
            End If
            frmWBXJ.adoWb.Close
            frmWBXJ.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            If frmWBXJ.adoWb.RecordCount > 1 Then
                frmWBXJ.dtgWb.FixedRows = 0
                frmWBXJ.dtgWb.MergeCol(1) = True
                frmWBXJ.dtgWb.MergeCol(2) = True
                frmWBXJ.dtgWb.MergeCol(3) = True
                frmWBXJ.dtgWb.MergeCells = 3
                frmWBXJ.dtgWb.FixedRows = 1
            End If
            Set frmWBXJ.dtgWb.DataSource = frmWBXJ.adoWb
            '例检表
            If frmWBXJ.comPb.Text <> "" Then '如果为旧版本,则显示年保和例检条款.
                tt = "select * from xunJIaWbView where wbx='例检' and bid=" & Val(frmWBXJ.lblBid.Caption)
            Else
                tt = "select * from xunJIaWbView where wbx='例检' and bid=0"
            End If
            frmWBXJ.adoLj.Close
            frmWBXJ.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            If frmWBXJ.adoLj.RecordCount > 1 Then
                frmWBXJ.dtgLj.FixedRows = 0
                frmWBXJ.dtgLj.MergeCol(1) = True
                frmWBXJ.dtgLj.MergeCol(2) = True
                frmWBXJ.dtgLj.MergeCol(3) = True
                frmWBXJ.dtgLj.MergeCells = 3
                frmWBXJ.dtgLj.FixedRows = 1
            End If
            Set frmWBXJ.dtgLj.DataSource = frmWBXJ.adoLj
            frmWBXJ.frmDx.Visible = False
            frmWBXJ.frmNb.Visible = True
            frmWBXJ.frmTime.Visible = True

            frmWBXJ.cmdD.Visible = True
            frmWBXJ.cmdJi.Visible = True
            frmWBXJ.tabGc.TabVisible(2) = False
            frmWBXJ.tabGc.TabVisible(0) = True
            frmWBXJ.tabGc.TabVisible(1) = True
            frmWBXJ.tabGc.Tab = 0
            
            
            
            If frmWBXJ.comPb.Text = "" Then '如果为新版的机组信息,则显示机组信息表
                frmWBXJ.frmNew.Visible = True
                tt = "select jzpb as 机组品牌,jzxh as 机组型号,sl as 数量,wid from wbjb where bid=" & Val(frmWBXJ.lblBid.Caption)
                Set frmWBXJ.adoA = CreateObject("adodb.recordset")
                frmWBXJ.adoA.Close
                frmWBXJ.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
                Set frmWBXJ.dtgA.DataSource = frmWBXJ.adoA
            Else
                frmWBXJ.frmNew.Visible = False
            End If
            frmWBXJ.cmdTK.Visible = True
         Else '大修

            frmWBXJ.frmDx.Visible = True
            frmWBXJ.frmNb.Visible = False
            frmWBXJ.frmTime.Visible = False

            frmWBXJ.cmdD.Visible = False
            frmWBXJ.cmdJi.Visible = Fal
            frmWBXJ.tabGc.TabVisible(2) = True
            frmWBXJ.tabGc.TabVisible(0) = False
            frmWBXJ.tabGc.TabVisible(1) = False
            frmWBXJ.tabGc.Tab = 2
            If frmWBXJ.comPb.Text = "" Then '如果为新版的机组信息,则显示机组信息表
                frmWBXJ.frmNew.Visible = True
                tt = "select jzpb as 机组品牌,jzxh as 机组型号,sl as 数量,wid from wbjb where bid=" & Val(frmWBXJ.lblBid.Caption)
                Set frmWBXJ.adoA = CreateObject("adodb.recordset")
                frmWBXJ.adoA.Close
                frmWBXJ.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
                Set frmWBXJ.dtgA.DataSource = frmWBXJ.adoA
            Else
                frmWBXJ.frmNew.Visible = False
            End If
            frmWBXJ.cmdTK.Value = False
         End If
        frmWBXJ.cmdMod.Enabled = True
        frmWBXJ.cmdSave.Enabled = False
        frmWBXJ.cmdD.Visible = False

        frmWBXJ.cmdJi.Visible = False
                frmWBXJ.frmQm.Visible = False
            frmWBXJ.lblTX.Visible = False
        If frmWBXJ.lblYwy.Caption = mod1.DName Or (frmWBXJ.lblBM.Caption = mod1.Bm And mod1.BmJl = True) Or mod1.DName = "宋晓炯" Or mod1.DName = "宋晓炯1" Or mod1.DName = "倪旭" Then
            frmWBXJ.cmdCG.Visible = True
'            If frmWBXJ.lblYwy.Caption = mod1.DName Then
'                frmWBXJ.cmdCong.Visible = True
'            Else
'                frmWBXJ.cmdCong.Visible = False
'            End If
        Else
            frmWBXJ.cmdCG.Visible = False

            frmWBXJ.cmdBjd.Visible = False
        End If
        tt = "select lc from htping where hid=" & Val(frmWBXJ.lblHtbh.Caption)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        frmWBXJ.lblHLC.Caption = mod1.HTP.Fields("lc").Value

        Call modBJD.OpenXJAN(LX)
''''                If Val(frmWBXJ.lblBid.Caption) >= 6794 Then
''''                    frmWBXJ.lblQM(2).Caption = "商务支持"
''''                End If
'''        '如果业务员本人打开,而且此单还未生成报价单,则允许重新询价
'''        If frmWBXJ.lblYwy.Caption = mod1.DName And Val(frmWBXJ.lblBaoid.Caption) = 0 And Val(frmWBXJ.lblLc.Caption) > 0 Then
'''            frmWBXJ.cmdCong.Visible = True
'''        End If



    End If
    
ElseIf ZL = "购销" Or ZL = "配件" Or ZL = "产品" Or ZL = "配件询价单" Or ZL = "零配件" Or ZL = "三菱" Or ZL = "松下" Or ZL = "勤达富" Or ZL = "德图" Or ZL = "非代理产品" Then '购销
    LX = False
    If mod1.DName = "谢雪梅" Or Bid > 10058 Then
        'frmGXBj.frmSD.Visible = True
        frmGXBj.frmCg.Top = 4740
        frmGXBj.dtgNew.Visible = True

        frmGXBj.dtgP.Visible = True
        frmGXBj.cmdGy.Visible = True
    Else
        'frmGXBj.frmSD.Visible = False
        frmGXBj.frmCg.Top = 7620
        frmGXBj.dtgNew.Visible = False

        frmGXBj.dtgP.Visible = False
        frmGXBj.cmdGy.Visible = False
    End If
        
    frmGXBj.JZ = mod1.HTP.Fields("jz").Value
    frmGXBj.lblZl.Caption = mod1.HTP.Fields("Zl").Value
    frmGXBj.comXmmc.Text = mod1.HTP.Fields("xmmc").Value
    frmGXBj.comXmmc.Tag = mod1.HTP.Fields("xid").Value
    frmGXBj.lblBid.Caption = mod1.HTP.Fields("bid").Value
    'bid = frmGXBj.lblBid.Caption
    frmGXBj.lblOid.Caption = mod1.HTP.Fields("oid").Value
    frmGXBj.lblLc.Caption = mod1.HTP.Fields("lc").Value
    frmGXBj.lblLcRen.Caption = mod1.HTP.Fields("lcren").Value
    frmGXBj.lblLcUid.Caption = mod1.HTP.Fields("lcuid").Value
    frmGXBj.lblFwid.Caption = mod1.HTP.Fields("fwid").Value
    frmGXBj.lblNlb.Caption = mod1.HTP.Fields("nlb").Value
    frmGXBj.lblLcou.Caption = mod1.HTP.Fields("lcou").Value
    frmGXBj.lblBaoId.Caption = mod1.HTP.Fields("baoid").Value
    frmGXBj.lblBh.Caption = "XJD" & mod1.HTP.Fields("bid").Value
    frmGXBj.lblPwf.Caption = mod1.HTP.Fields("pwf").Value
        'frmGXBj.txtHg.Text = mod1.HTP.Fields("hg").Value
        frmGXBj.LBLhG.Caption = mod1.HTP.Fields("hg").Value
        'frmGXBj.txtYhg.Text = mod1.HTP.Fields("yhg").Value
        frmGXBj.LBLyHG.Caption = mod1.HTP.Fields("jhg").Value
        frmGXBj.lblWhg.Caption = mod1.HTP.Fields("jhg").Value
    frmGXBj.txtBz.Text = mod1.HTP.Fields("bz").Value
    frmGXBj.lblYwy.Caption = mod1.HTP.Fields("ywy").Value
    frmGXBj.lblUid.Caption = mod1.HTP.Fields("uid").Value
    frmGXBj.ZF = mod1.HTP.Fields("zf").Value
    frmGXBj.txtHtbh.Text = mod1.HTP.Fields("htbh").Value
    frmGXBj.txtYf.Text = mod1.HTP.Fields("yf").Value
    frmGXBj.txtADR.Text = mod1.HTP.Fields("yfadr").Value
    frmGXBj.lblHtbh.Caption = mod1.HTP.Fields("htbh").Value
    frmGXBj.txtFj.Text = mod1.HTP.Fields("fjhg").Value
    frmGXBj.lblTX.Caption = "现在流程跳至：" & frmGXBj.lblLcRen.Caption
    If Val(frmGXBj.lblLcRen.Caption) = 100 Then
        frmGXBj.lblTX.Caption = "流程完毕"
    End If
    If mod1.HTP.Fields("chf").Value = True And frmGXBj.lblLc.Caption > 2 Then
        frmGXBj.lblZ.Visible = True
        frmGXBj.lblZT.Visible = True
        frmGXBj.lblZT.Caption = mod1.HTP.Fields("chdate").Value

    End If
        frmGXBj.lblCfwid.Caption = mod1.HTP.Fields("cfwid").Value
    tt = "select qy,bm from RenYuan where userName='" & frmGXBj.lblYwy.Caption & "' and userid='" & frmGXBj.lblUid.Caption & "'"
    mod1.HTT.Close
    mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If IsNull(mod1.HTT.RecordCount) = True Then
        MsgBox ("连接超时,将退出!")
        End
    End If
    frmGXBj.lblBM.Caption = mod1.HTT.Fields("bm").Value
    frmGXBj.lblQy.Caption = mod1.HTT.Fields("qy").Value
    
    tt = "select * from xunJIamxView where bid=" & Val(frmGXBj.lblBid.Caption)
    frmGXBj.adoGx.Close
    frmGXBj.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    If IsNull(frmGXBj.adoGx.RecordCount) = True Then
        MsgBox ("连接超时,将退出!")
        End
    End If
    Set frmGXBj.dtgMa.DataSource = frmGXBj.adoGx
frmGXBj.dtgN.Rows = frmGXBj.dtgMa.Rows + 20
frmGXBj.dtgN.Cols = frmGXBj.dtgMa.Cols
    
    For oo = 0 To frmGXBj.dtgMa.Cols - 1
        frmGXBj.dtgMa.Col = oo
        frmGXBj.dtgMa.Row = 0
        If frmGXBj.dtgMa.Text = "机组型号" Or frmGXBj.dtgMa.Text = "零件编号" Or frmGXBj.dtgMa.Text = "零件名称" Then
            
            frmGXBj.dtgMa.ColWidth(oo) = 2000

        End If

        If frmGXBj.dtgMa.Text = "到货期" Or frmGXBj.dtgMa.Text = "报价有效期" Then
            frmGXBj.dtgMa.ColWidth(oo) = 1500
        End If
        If frmGXBj.dtgMa.Text = "压缩机型号 " Or frmGXBj.dtgMa.Text = "出厂编号" Or frmGXBj.dtgMa.Text = "机组序列号" Or frmGXBj.dtgMa.Text = "品牌产地" Or frmGXBj.dtgMa.Text = "市场价" Or _
        frmGXBj.dtgMa.Text = "bid" Or frmGXBj.dtgMa.Text = "Lid" Or frmGXBj.dtgMa.Text = "gyId" Or frmGXBj.dtgMa.Text = "gyBZ" Or frmGXBj.dtgMa.Text = "品种" Then
            frmGXBj.dtgMa.ColWidth(oo) = 0
        End If
            If frmGXBj.lblYwy = "谢雪梅" Or Bid > 10058 Then
                If frmGXBj.dtgMa.Text = "压缩机型号" Then
                    frmGXBj.dtgMa.Text = "单位"
                    frmGXBj.dtgMa.ColWidth(oo) = 500
                ElseIf frmGXBj.dtgMa.Text = "机组型号" Then
                    frmGXBj.dtgMa.ColWidth(oo) = 1500
                ElseIf frmGXBj.dtgMa.Text = "零件编号" Then
                    frmGXBj.dtgMa.ColWidth(oo) = 1000
                    frmGXBj.dtgMa.Text = "货品编码"
                ElseIf frmGXBj.dtgMa.Text = "品牌产地" Then
                    frmGXBj.dtgMa.Text = "规格"
                    frmGXBj.dtgMa.ColWidth(oo) = 2500
                ElseIf frmGXBj.dtgMa.Text = "零件名称" Then

                    frmGXBj.dtgMa.Text = "货品名称"
                ElseIf frmGXBj.dtgMa.Text = "质保期" Then
                    frmGXBj.dtgMa.ColWidth(oo) = 1000
                End If
                
            End If
        If lblUid.Caption = mod1.DHid Then  '业务员，只显示基准价
            If frmGXBj.dtgMa.Text = "成本单价" Or frmGXBj.dtgMa.Text = "合计" Or frmGXBj.dtgMa.Text = "外包单价" Or frmGXBj.dtgMa.Text = "外包合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 0
            End If
            If frmGXBj.dtgMa.Text = "基准单价" Or frmGXBj.dtgMa.Text = "基准合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 1000
            End If
        ElseIf mod1.Bm = "配送中心" And mod1.BmJl = False Then
            If frmGXBj.dtgMa.Text = "成本单价" Or frmGXBj.dtgMa.Text = "合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 1000
            End If
            If frmGXBj.dtgMa.Text = "基准单价" Or frmGXBj.dtgMa.Text = "基准合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 0
            End If
        ElseIf mod1.Bm = "配送中心" And mod1.BmJl = True Then
            If frmGXBj.dtgMa.Text = "成本单价" Or frmGXBj.dtgMa.Text = "合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 1000
            End If
            If frmGXBj.dtgMa.Text = "基准单价" Or frmGXBj.dtgMa.Text = "基准合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 1000
            End If
        ElseIf mod1.Bm = "商务部" Then '其他人员都能看到
            If frmGXBj.dtgMa.Text = "成本单价" Or frmGXBj.dtgMa.Text = "合计" Or frmGXBj.dtgMa.Text = "基准单价" Or frmGXBj.dtgMa.Text = "基准合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 1000
            End If
        Else
            If frmGXBj.dtgMa.Text = "成本单价" Or frmGXBj.dtgMa.Text = "合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 0
            End If
            If frmGXBj.dtgMa.Text = "基准单价" Or frmGXBj.dtgMa.Text = "基准合计" Then
                frmGXBj.dtgMa.ColWidth(oo) = 1000
            End If
        
        End If
    Next
        Set frmGXBj.dtgMa.DataSource = Nothing
        
        
    '显示商务支持添加的产品（变色）
    Dim jj As Integer
    For oo = 1 To frmGXBj.dtgMa.Rows + 1
        frmGXBj.dtgMa.Col = 28
        frmGXBj.dtgMa.Row = oo
        frmGXBj.dtgN.Row = oo
        If frmGXBj.dtgMa.Text = "True" Then
            For jj = 1 To 25
                frmGXBj.dtgMa.Col = jj
                frmGXBj.dtgMa.CellForeColor = &HFF0000
            Next
        End If
        For jj = 1 To 25
            frmGXBj.dtgMa.Col = jj
            frmGXBj.dtgN.Col = jj
            frmGXBj.dtgN.Text = frmGXBj.dtgMa.Text
        Next
    Next
''''''    If frmGXBj.adoGx.RecordCount > 0 Then
''''''        frmGXBj.dtgMa.FixedRows = 0
''''''        frmGXBj.dtgMa.MergeCol(1) = True
''''''        frmGXBj.dtgMa.MergeCol(2) = True
''''''        frmGXBj.dtgMa.MergeCol(10) = True
''''''        frmGXBj.dtgMa.MergeCol(14) = True
''''''        frmGXBj.dtgMa.MergeCol(15) = True
''''''        frmGXBj.dtgMa.MergeCells = 3
''''''        frmGXBj.dtgMa.FixedRows = 1
''''''    End If
    frmGXBj.cmdSave.Enabled = False
    frmGXBj.cmdMod.Enabled = True

''    If mod1.DName = frmGXBj.lblYwy.Caption Then
''        frmGXBj.cmdCong.Visible = True
''    End If
    If mod1.Bm = "零件事业部" Then
        frmGXBj.OPTN.Value = True
        frmGXBj.txtHg.Text = frmGXBj.LBLhG.Caption
        frmGXBj.txtYhg.Text = frmGXBj.LBLyHG.Caption
        frmGXBj.cmdHT.Visible = False
        frmGXBj.optW.Enabled = False
    Else
        frmGXBj.optW.Value = True
        frmGXBj.txtHg.Text = frmGXBj.lblWhg.Caption
        frmGXBj.txtYhg.Text = frmGXBj.lblWhg.Caption
        frmGXBj.cmdHT.Visible = True
        frmGXBj.optW.Enabled = True
    End If
    tt = "select lc,prf from htping where hid=" & Val(frmGXBj.lblHtbh.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    frmGXBj.lblHLC.Caption = mod1.HTP.Fields("lc").Value
    frmGXBj.lblHLC.ToolTipText = mod1.HTP.Fields("prf").Value '纯配件还是商务支持
    Call modBJD.OpenXJAN(LX)
    frmGXBj.cmdWb.Visible = False
''''    If Val(frmGXBj.lblLc.Caption) = 2 And frmGXBj.cmdQm(1).Caption <> "" And frmGXBj.cmdQm(1).Caption <> "" And Val(frmGXBj.lblHLC.ToolTipText) = 2 Then '让商务支持能够签字
''''        frmGXBj.lblQM(1).Caption = "商务支持"
''''        frmGXBj.cmdQm(1).Caption = ""
''''        frmGXBj.lblTm(1).Caption = ""
''''    End If
    Call frmGXBj.QMBound(Bid)
End If









End Sub

Public Sub BaoJDBound(BaoId As Long, ZL As String)   '
Dim tt As String
Dim xmFy As Single '已发生的项目费用
On Error Resume Next
frmWbxjB.frmDx.Visible = False
frmWbxjB.frmNb.Visible = False
frmWbxjB.frmTime.Visible = False


tt = "select top 1 * from baoJiaD where baoid=" & BaoId
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
ZL = mod1.HTP.Fields("zl").Value
'先计算已发生的项目费用
If IsNull(mod1.HTP.Fields("htbh").Value) = True Then
    tt = "select (xmfy-ygfy) as xmfy from xmzl where xid=" & mod1.HTP.Fields("xid").Value
Else
    tt = "select ygfy as xmfy from xmzl where xid=" & mod1.HTP.Fields("xid").Value
End If
mod1.HTT.Close
mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
xmFy = mod1.HTT.Fields("xmfy").Value

tt = "select sum(xg) from fybx "
If ZL = "维保" Then '维保
    frmWbxjB.frmJz.Visible = True
    frmWbxjB.frmNb.Visible = True
    frmWbxjB.lblZl.Caption = mod1.HTP.Fields("zl").Value
    frmWbxjB.comXmmc.Text = mod1.HTP.Fields("xmmc").Value
    frmWbxjB.comXmmc.Tag = mod1.HTP.Fields("xid").Value
    frmWbxjB.comKhmc.Text = mod1.HTP.Fields("khmc").Value
    frmWbxjB.comKhmc.ToolTipText = mod1.HTP.Fields("khdh").Value
    frmWbxjB.lblBid.Caption = mod1.HTP.Fields("bid").Value
    frmWbxjB.comZu.Text = mod1.HTP.Fields("zh").Value
    frmWbxjB.txtZu.Text = mod1.HTP.Fields("zName").Value
    frmWbxjB.comPb.Text = mod1.HTP.Fields("jzpb").Value
    frmWbxjB.comXh.Text = mod1.HTP.Fields("jzxh").Value
    frmWbxjB.txtSL.Text = mod1.HTP.Fields("sL").Value
    frmWbxjB.lblOid.Caption = mod1.HTP.Fields("oid").Value
    frmWbxjB.txtZT.Text = mod1.HTP.Fields("ZTime").Value
    frmWbxjB.txtHg.Text = mod1.HTP.Fields("bHG").Value
    frmWbxjB.txtYhg.Text = mod1.HTP.Fields("yhg").Value
    frmWbxjB.chkBa.Value = Abs(CInt(mod1.HTP.Fields("ta").Value))
    frmWbxjB.chkBb.Value = Abs(CInt(mod1.HTP.Fields("tb").Value))
    frmWbxjB.chkBc.Value = Abs(CInt(mod1.HTP.Fields("tc").Value))
    frmWbxjB.lblYwy.Caption = mod1.HTP.Fields("ywy").Value
    frmWbxjB.lblUid.Caption = mod1.HTP.Fields("uid").Value
    frmWbxjB.lblBaoId.Caption = mod1.HTP.Fields("baoid").Value
    frmWbxjB.lblBh.Caption = mod1.HTP.Fields("baoid").Value
    frmWbxjB.txtFbje.Text = mod1.HTP.Fields("fbje").Value
    frmWbxjB.txtFbnr.Text = mod1.HTP.Fields("fbnr").Value
    
    frmWbxjB.lblLc.Caption = mod1.HTP.Fields("Lc").Value
    frmWbxjB.lblLcRen.Caption = mod1.HTP.Fields("LcRen").Value
    frmWbxjB.lblLcUid.Caption = mod1.HTP.Fields("LcUid").Value
    frmWbxjB.lblFwid.Caption = mod1.HTP.Fields("Fwid").Value
    frmWbxjB.lblNlb.Caption = mod1.HTP.Fields("Nlb").Value
    frmWbxjB.lblLcou.Caption = mod1.HTP.Fields("Lcou").Value
    frmWbxjB.txtRgf.Text = mod1.HTP.Fields("rgf").Value
    frmWbxjB.txtClf.Text = mod1.HTP.Fields("clf").Value
    frmWbxjB.txtClcb.Text = mod1.HTP.Fields("clcb").Value
    frmWbxjB.txtYJ.Text = mod1.HTP.Fields("yj").Value
    'frmWbxjB.txtCb.Text = mod1.HTP.Fields("rgf").Value + mod1.HTP.Fields("clf").Value + mod1.HTP.Fields("clcb").Value + xmFy
    frmWbxjB.txtCb.Text = mod1.HTP.Fields("hg").Value
    frmWbxjB.txtXm1.Text = xmFy
    frmWbxjB.txtXm2.Text = mod1.HTP.Fields("ylxm").Value
    
    frmWbxjB.txtMon.Text = mod1.HTP.Fields("mon").Value
    frmWbxjB.txtWc.Text = mod1.HTP.Fields("wc").Value
    frmWbxjB.txtXc.Text = mod1.HTP.Fields("xc").Value
    frmWbxjB.txtYf.Text = mod1.HTP.Fields("yf").Value
    frmWbxjB.lblHtbh.Caption = mod1.HTP.Fields("htbh").Value
    frmWbxjB.txtF.Text = mod1.HTP.Fields("htqy").Value
    frmWbxjB.txtL.Text = mod1.HTP.Fields("htqy1").Value
    frmWbxjB.txtTcBe.Text = mod1.HTP.Fields("tcbe").Value
    frmWbxjB.txtBz.Text = mod1.HTP.Fields("bz").Value
    If mod1.HTP.Fields("fpLX").Value = "增值发票" Then
        frmWbxjB.optLa.Value = True
    ElseIf mod1.HTP.Fields("fpLX").Value = "商业发票" Then
        frmWbxjB.optLb.Value = True
    ElseIf mod1.HTP.Fields("fpLX").Value = "服务发票" Then
        frmWbxjB.optLc.Value = True
    End If
    If frmWbxjB.comPb.Text = "" Then
        tt = "select jzpb as 机组品牌,jzxh as 机组型号,sl as 数量,jxId from wbjb where baoid=" & Val(frmWbxjB.lblBaoId.Caption)
        Set frmWbxjB.adoA = CreateObject("adodb.recordset")
        frmWbxjB.adoA.Close
        frmWbxjB.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmWbxjB.dtgA.DataSource = frmWbxjB.adoA
        frmWbxjB.dtgA.Visible = True
        frmWbxjB.cmdTK.Visible = True
    Else
        frmWbxjB.dtgA.Visible = False
        '年保表
        tt = "select * from xunJIaWbView where wbx='年保' and bid=" & Val(frmWbxjB.lblBid.Caption)
        frmWbxjB.adoWb.Close
        frmWbxjB.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmWbxjB.dtgWb.DataSource = frmWbxjB.adoWb
        frmWbxjB.dtgWb.FixedRows = 0
        frmWbxjB.dtgWb.MergeCol(1) = True
        frmWbxjB.dtgWb.MergeCol(2) = True
        frmWbxjB.dtgWb.MergeCol(3) = True
        frmWbxjB.dtgWb.MergeCells = 3
        frmWbxjB.dtgWb.FixedRows = 1
        '例检表
        tt = "select * from xunJIaWbView where wbx='例检' and bid=" & Val(frmWbxjB.lblBid.Caption)
        frmWbxjB.adoLj.Close
        frmWbxjB.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmWbxjB.dtgLj.DataSource = frmWbxjB.adoLj
        frmWbxjB.dtgLj.FixedRows = 0
        frmWbxjB.dtgLj.MergeCol(1) = True
        frmWbxjB.dtgLj.MergeCol(2) = True
        frmWbxjB.dtgLj.MergeCol(3) = True
        frmWbxjB.dtgLj.MergeCells = 3
        frmWbxjB.dtgLj.FixedRows = 1
        frmWbxjB.cmdTK.Visible = False
    End If
    
    '显示产品列表
    tt = "select * from BaoJiaMxView where baoid=" & Val(frmWbxjB.lblBaoId.Caption) & " order by lid"
    frmWbxjB.adoBx.Close
    frmWbxjB.adoBx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbxjB.dtgBao.DataSource = frmWbxjB.adoBx
    frmWbxjB.dtgBao.FixedRows = 0
    frmWbxjB.dtgBao.MergeCol(1) = True
    frmWbxjB.dtgBao.MergeCol(2) = True
    frmWbxjB.dtgBao.MergeCol(10) = True
    frmWbxjB.dtgBao.MergeCol(14) = True
    frmWbxjB.dtgBao.MergeCells = 3
    frmWbxjB.dtgBao.FixedRows = 1
    '显示成本表
    tt = "select * from xunJiaMxView where bid=0"
    frmWbxjB.adoGx.Close
    frmWbxjB.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbxjB.dtgMa.DataSource = frmWbxjB.adoGx
    
    If mod1.HTP.Fields("pwf").Value = True Then '如果评审完毕,则打印报价单
        frmWbxjB.cmdPrint.Visible = True
        frmWbxjB.cmdHT.Visible = True
    Else
        frmWbxjB.cmdPrint.Visible = False
        frmWbxjB.cmdHT.Visible = False
    End If
    
    '打开佣金表
    tt = "select yED as 收款额度,YingFu as 支付金额,yid from byj where baoid=" & frmWbxjB.lblBaoId.Caption & " order by yid"
    frmWbxjB.adoYj.Close
    frmWbxjB.adoYj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbxjB.dtgYJ.DataSource = frmWbxjB.adoYj
    frmWbxjB.frmYm.Visible = False
    '打开付款表
    tt = "select * from baoFk where baoid=" & frmWbxjB.lblBaoId.Caption & " order by fid"
    frmWbxjB.adoFk.Close
    frmWbxjB.adoFk.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbxjB.dtgFk.DataSource = frmWbxjB.adoFk
    frmWbxjB.frmFF.Visible = False
    
'    frmWbxjB.lblYj.Visible = False
'    frmWbxjB.txtYj.Visible = False
    frmWbxjB.frmTime.Visible = True
    frmWbxjB.tabGc.TabVisible(0) = True
    frmWbxjB.tabGc.TabVisible(1) = True
    frmWbxjB.tabGc.TabVisible(2) = False
    frmWbxjB.tabGc.TabVisible(3) = True
    frmWbxjB.cmdCong.Visible = False
    Call modBJD.OpenBJAN(1)
    Call modBJD.wbxjbLocked
    frmWbxjB.Visible = True
    frmWbxjB.cmdSave.Enabled = False
    frmWbxjB.frmYj.Visible = False
        tt = "select qy,bm from RenYuan where userName='" & frmWbxjB.lblYwy.Caption & "' and userid='" & frmWbxjB.lblUid.Caption & "'"
        mod1.HTT.Close
        mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        frmWbxjB.lblBM.Caption = mod1.HTT.Fields("bm").Value
        frmWbxjB.lblQy.Caption = mod1.HTT.Fields("qy").Value
ElseIf ZL = "大修" Or ZL = "工程分包" Then
    frmWbxjB.tabGc.TabCaption(2) = mod1.HTP.Fields("zl").Value
    frmWbxjB.Visible = True
    frmWbxjB.frmJz.Visible = False
    frmWbxjB.frmDx.Visible = True
    frmWbxjB.lblZl.Caption = mod1.HTP.Fields("zl").Value
    frmWbxjB.comXmmc.Text = mod1.HTP.Fields("xmmc").Value
    frmWbxjB.comXmmc.Tag = mod1.HTP.Fields("xid").Value
    frmWbxjB.comKhmc.Text = mod1.HTP.Fields("khmc").Value
    frmWbxjB.comKhmc.ToolTipText = mod1.HTP.Fields("khdh").Value
    frmWbxjB.lblBid.Caption = mod1.HTP.Fields("bid").Value
    frmWbxjB.comZu.Text = mod1.HTP.Fields("zh").Value
    frmWbxjB.txtZu.Text = mod1.HTP.Fields("zName").Value
    frmWbxjB.comPb.Text = mod1.HTP.Fields("jzpb").Value
    frmWbxjB.comXh.Text = mod1.HTP.Fields("jzxh").Value
    frmWbxjB.txtSL.Text = mod1.HTP.Fields("sL").Value
    frmWbxjB.lblOid.Caption = mod1.HTP.Fields("oid").Value
    frmWbxjB.txtZT.Text = mod1.HTP.Fields("ZTime").Value
    frmWbxjB.txtHg.Text = mod1.HTP.Fields("bHG").Value
    frmWbxjB.txtYhg.Text = mod1.HTP.Fields("yhg").Value
    frmWbxjB.chkBa.Value = Abs(CInt(mod1.HTP.Fields("ta").Value))
    frmWbxjB.chkBb.Value = Abs(CInt(mod1.HTP.Fields("tb").Value))
    frmWbxjB.chkBc.Value = Abs(CInt(mod1.HTP.Fields("tc").Value))
    frmWbxjB.lblYwy.Caption = mod1.HTP.Fields("ywy").Value
    frmWbxjB.lblUid.Caption = mod1.HTP.Fields("uid").Value
    frmWbxjB.lblBaoId.Caption = mod1.HTP.Fields("baoid").Value
    frmWbxjB.lblBh.Caption = mod1.HTP.Fields("baoid").Value
    frmWbxjB.txtFbje.Text = mod1.HTP.Fields("fbje").Value
    frmWbxjB.txtFbnr.Text = mod1.HTP.Fields("fbnr").Value
    
    frmWbxjB.lblLc.Caption = mod1.HTP.Fields("Lc").Value
    frmWbxjB.lblLcRen.Caption = mod1.HTP.Fields("LcRen").Value
    frmWbxjB.lblLcUid.Caption = mod1.HTP.Fields("LcUid").Value
    frmWbxjB.lblFwid.Caption = mod1.HTP.Fields("Fwid").Value
    frmWbxjB.lblNlb.Caption = mod1.HTP.Fields("Nlb").Value
    frmWbxjB.lblLcou.Caption = mod1.HTP.Fields("Lcou").Value
    frmWbxjB.txtRgf.Text = mod1.HTP.Fields("rgf").Value
    frmWbxjB.txtClf.Text = mod1.HTP.Fields("clf").Value
    frmWbxjB.txtClcb.Text = mod1.HTP.Fields("clcb").Value
    frmWbxjB.txtYJ.Text = mod1.HTP.Fields("yj").Value
    frmWbxjB.txtCb.Text = mod1.HTP.Fields("hg").Value
    frmWbxjB.txtYf.Text = mod1.HTP.Fields("yf").Value
    frmWbxjB.txtXm1.Text = xmFy
    frmWbxjB.txtXm2.Text = mod1.HTP.Fields("ylxm").Value
    frmWbxjB.txtMon.Text = mod1.HTP.Fields("mon").Value
    frmWbxjB.txtDxnr.Text = mod1.HTP.Fields("dxnr").Value
    frmWbxjB.txtWc.Text = mod1.HTP.Fields("wc").Value
    frmWbxjB.txtXc.Text = mod1.HTP.Fields("xc").Value
    frmWbxjB.lblHtbh.Caption = mod1.HTP.Fields("htbh").Value
    frmWbxjB.txtTcBe.Text = mod1.HTP.Fields("tcbe").Value
    frmWbxjB.txtBz.Text = mod1.HTP.Fields("bz").Value
    If mod1.HTP.Fields("fpLX").Value = "增值发票" Then
        frmWbxjB.optLa.Value = True
    ElseIf mod1.HTP.Fields("fpLX").Value = "商业发票" Then
        frmWbxjB.optLb.Value = True
    ElseIf mod1.HTP.Fields("fpLX").Value = "服务发票" Then
        frmWbxjB.optLc.Value = True
    End If
    frmWbxjB.cmdTK.Visible = False
    If frmWbxjB.comPb.Text = "" Then
        tt = "select jzpb as 机组品牌,jzxh as 机组型号,sl as 数量,jxId from wbjb where baoid=" & Val(frmWbxjB.lblBaoId.Caption)
        Set frmWbxjB.adoA = CreateObject("adodb.recordset")
        frmWbxjB.adoA.Close
        frmWbxjB.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmWBXJ.dtgA.DataSource = frmWBXJ.adoA
        frmWbxjB.dtgA.Visible = True
        frmWbxjB.cmdTK.Visible = True
    Else
        frmWbxjB.cmdTK.Visible = False
        
    End If

    '显示产品列表
    tt = "select * from BaoJiaMxView where baoid=" & Val(frmWbxjB.lblBaoId.Caption) & " order by lid"
    frmWbxjB.adoBx.Close
    frmWbxjB.adoBx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbxjB.dtgBao.DataSource = frmWbxjB.adoBx
    frmWbxjB.dtgBao.FixedRows = 0
    frmWbxjB.dtgBao.MergeCol(1) = True
    frmWbxjB.dtgBao.MergeCol(2) = True
    frmWbxjB.dtgBao.MergeCol(10) = True
    frmWbxjB.dtgBao.MergeCol(14) = True
    frmWbxjB.dtgBao.MergeCells = 3
    frmWbxjB.dtgBao.FixedRows = 1
    '显示成本表
    tt = "select * from xunJiaMxView where bid=0"
    frmWbxjB.adoGx.Close
    frmWbxjB.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbxjB.dtgMa.DataSource = frmWbxjB.adoGx
    If mod1.HTP.Fields("pwf").Value = True Then '如果评审完毕,则打印报价单
        frmWbxjB.cmdPrint.Visible = True
        frmWbxjB.cmdHT.Visible = True
    Else
        frmWbxjB.cmdPrint.Visible = False
        frmWbxjB.cmdHT.Visible = False
    End If
    '打开佣金表
    tt = "select yED as 收款额度,YingFu as 支付金额,yid from byj  where baoid=" & frmWbxjB.lblBaoId.Caption & " order by yid"
    frmWbxjB.adoYj.Close
    frmWbxjB.adoYj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbxjB.dtgYJ.DataSource = frmWbxjB.adoYj
    '打开付款表
    tt = "select * from baoFk where baoid=" & frmWbxjB.lblBaoId.Caption & " order by fid"
    frmWbxjB.adoFk.Close
    frmWbxjB.adoFk.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWbxjB.dtgFk.DataSource = frmWbxjB.adoFk
    frmWbxjB.frmFF.Visible = False
    frmWbxjB.frmYm.Visible = False
    frmWbxjB.tabGc.TabVisible(0) = False
    frmWbxjB.tabGc.TabVisible(1) = False
    frmWbxjB.tabGc.TabVisible(2) = True
    frmWbxjB.tabGc.TabVisible(3) = True
    frmWbxjB.tabGc.Tab = 2
    frmWbxjB.cmdCong.Visible = False
    frmWbxjB.cmdSave.Enabled = False
    Call modBJD.OpenBJAN(1)
    Call modBJD.wbxjbLocked
    frmWbxjB.frmYj.Visible = False
        tt = "select qy,bm from RenYuan where userName='" & frmWbxjB.lblYwy.Caption & "' and userid='" & frmWbxjB.lblUid.Caption & "'"
        mod1.HTT.Close
        mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        frmWbxjB.lblBM.Caption = mod1.HTT.Fields("bm").Value
        frmWbxjB.lblQy.Caption = mod1.HTT.Fields("qy").Value
ElseIf ZL = "购销" Then '购销
    frmGxbjB.Visible = True
    frmGxbjB.lblZl.Caption = mod1.HTP.Fields("zl").Value
    frmGxbjB.comXmmc.Text = mod1.HTP.Fields("xmmc").Value
    frmGxbjB.comXmmc.Tag = mod1.HTP.Fields("xid").Value
    frmGxbjB.comKhmc.Text = mod1.HTP.Fields("khmc").Value
    frmGxbjB.comKhmc.ToolTipText = mod1.HTP.Fields("khdh").Value
    frmGxbjB.lblBid.Caption = mod1.HTP.Fields("bid").Value
    frmGxbjB.lblOid.Caption = mod1.HTP.Fields("oid").Value
    frmGxbjB.txtHg.Text = mod1.HTP.Fields("bHG").Value
    frmGxbjB.txtYhg.Text = mod1.HTP.Fields("yhg").Value
    frmGxbjB.lblYwy.Caption = mod1.HTP.Fields("ywy").Value
    frmGxbjB.lblUid.Caption = mod1.HTP.Fields("uid").Value
    frmGxbjB.lblBaoId.Caption = mod1.HTP.Fields("baoid").Value
    frmGxbjB.lblBh.Caption = mod1.HTP.Fields("baoid").Value
    frmGxbjB.txtFbje.Text = mod1.HTP.Fields("fbje").Value
    frmGxbjB.txtFbnr.Text = mod1.HTP.Fields("fbnr").Value
    frmGxbjB.lblLc.Caption = mod1.HTP.Fields("Lc").Value
    frmGxbjB.lblLcRen.Caption = mod1.HTP.Fields("LcRen").Value
    frmGxbjB.lblLcUid.Caption = mod1.HTP.Fields("LcUid").Value
    
    frmGxbjB.lblFwid.Caption = mod1.HTP.Fields("Fwid").Value
    frmGxbjB.lblNlb.Caption = mod1.HTP.Fields("Nlb").Value
    frmGxbjB.lblLcou.Caption = mod1.HTP.Fields("Lcou").Value
    
    frmGxbjB.txtXm1.Text = xmFy
    frmGxbjB.txtXm.Text = mod1.HTP.Fields("ylxm").Value
    'frmGxbjB.txtXm2.Text = mod1.HTP.Fields("ylxm").Value
    frmGxbjB.txtClcb.Text = mod1.HTP.Fields("clcb").Value
    frmGxbjB.txtYJ.Text = mod1.HTP.Fields("yj").Value
    frmGxbjB.txtYf.Text = mod1.HTP.Fields("yf").Value
    frmGxbjB.txtCb.Text = mod1.HTP.Fields("hg").Value
    frmGxbjB.lblHtbh.Caption = mod1.HTP.Fields("htbh").Value
    frmGxbjB.txtTcBe.Text = mod1.HTP.Fields("tcbe").Value
    frmGxbjB.txtBz.Text = mod1.HTP.Fields("bz").Value
    
    If mod1.HTP.Fields("fpLX").Value = "增值发票" Then
        frmGxbjB.optLa.Value = True
    ElseIf mod1.HTP.Fields("fpLX").Value = "商业发票" Then
        frmGxbjB.optLb.Value = True
    ElseIf mod1.HTP.Fields("fpLX").Value = "服务发票" Then
        frmGxbjB.optLc.Value = True
    End If
    If mod1.HTP.Fields("pwf").Value = True Then '如果评审完毕,则生成报价单
        frmGxbjB.cmdPrint.Visible = True
        frmGxbjB.cmdHT.Visible = True
    Else
        frmGxbjB.cmdPrint.Visible = False
        frmGxbjB.cmdHT.Visible = False
    End If
    '显示产品列表
    tt = "select * from BaoJiaMxView where baoid=" & Val(frmGxbjB.lblBaoId.Caption) & " order by lid"
    frmGxbjB.adoBx.Close
    frmGxbjB.adoBx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGxbjB.dtgBao.DataSource = frmGxbjB.adoBx
    frmGxbjB.dtgBao.FixedRows = 0
    frmGxbjB.dtgBao.MergeCol(1) = True
    frmGxbjB.dtgBao.MergeCol(2) = True
    frmGxbjB.dtgBao.MergeCol(10) = True
    frmGxbjB.dtgBao.MergeCol(14) = True
    frmGxbjB.dtgBao.MergeCells = 3
    frmGxbjB.dtgBao.FixedRows = 1
    '显示成本表
    tt = "select * from xunJiaMxView where bid=0"
    frmGxbjB.adoGx.Close
    frmGxbjB.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGxbjB.dtgMa.DataSource = frmGxbjB.adoGx
    frmGxbjB.txtYhg.Locked = True
    '打开佣金表
    tt = "select yED as 收款额度,YingFu as 支付金额,yid from byj  where baoid=" & frmGxbjB.lblBaoId.Caption & " order by yid"
    frmGxbjB.adoYj.Close
    frmGxbjB.adoYj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGxbjB.dtgYJ.DataSource = frmGxbjB.adoYj
    frmGxbjB.frmYm.Visible = False
    '打开付款表
    tt = "select * from baoFk where baoid=" & frmGxbjB.lblBaoId.Caption & " order by fid"
    frmGxbjB.adoFk.Close
    frmGxbjB.adoFk.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGxbjB.dtgFk.DataSource = frmGxbjB.adoFk
    frmGxbjB.frmFF.Visible = False
    Call modBJD.OpenBJAN(0)
    Call modBJD.gxbjbLocked
    frmGxbjB.cmdSave.Enabled = False
    tt = "select qy,bm from RenYuan where userName='" & frmGxbjB.lblYwy.Caption & "' and userid='" & frmGxbjB.lblUid.Caption & "'"
    mod1.HTT.Close
    mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    frmGxbjB.lblBM.Caption = mod1.HTT.Fields("bm").Value
    frmGxbjB.lblQy.Caption = mod1.HTT.Fields("qy").Value
    '显示固定费用
    tt = "select lb as 费用类别,year(nd) as 年度,qdj as 单价,rl as 人数,xg as 小计,baoid,hid,gid from xmgd where baoid=" & Val(frmGxbjB.lblBaoId.Caption)
    frmGxbjB.adoGD.Close
    frmGxbjB.adoGD.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGxbjB.dtgGD.DataSource = frmGxbjB.adoGD
    tt = "select sum(xg) as xg from xmgd where baoid=" & Val(frmGxbjB.lblBaoId.Caption)
    frmGxbjB.adoHGD.Close
    frmGxbjB.adoHGD.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    frmGxbjB.txtGd.Text = frmGxbjB.adoHGD.Fields("xg").Value
    frmGxbjB.txtXm2.Text = Val(frmGxbjB.txtGd.Text) + Val(frmGxbjB.txtXm.Text)
End If
    

End Sub
Public Sub OpenXJAN(LX As Boolean)
Dim tt As String
Dim oo As Integer
On Error Resume Next
If LX = True Then   '维保
    For oo = 10 To 1 Step -1
        Unload frmWBXJ.cmdQm(oo)
        Unload frmWBXJ.lblQM(oo)
        Unload frmWBXJ.lblTm(oo)
    Next
    frmWBXJ.cmdQm(0).Caption = ""
    frmWBXJ.lblTm(0).Caption = ""
      tt = "qmrzOpen(" & mod1.BTZ & ",'" & frmWBXJ.lblBid.Caption & "')"
      Set mod1.HTP = CreateObject("adodb.recordset")
      mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    If IsNull(mod1.HTP.RecordCount) = True Then
        MsgBox ("连接超时,将退出!")
        End
    End If
      If mod1.HTP.RecordCount > 0 Then
         mod1.HTP.MoveFirst
         frmWBXJ.cmdQm(0).Visible = True
         frmWBXJ.lblQM(0).Visible = True
         frmWBXJ.lblTm(0).Visible = True
         frmWBXJ.lblQM(0).Caption = mod1.HTP.Fields("QLabel").Value
         If mod1.HTP.Fields("xf").Value = True Then
            frmWBXJ.cmdQm(0).Caption = mod1.HTP.Fields("Qren").Value
            frmWBXJ.lblTm(0).Caption = mod1.HTP.Fields("QRQ").Value
         End If
         frmWBXJ.cmdQm(0).Tag = mod1.HTP.Fields("zid").Value
         mod1.HTP.MoveNext
         For oo = 1 To mod1.HTP.RecordCount - 1
           Load frmWBXJ.lblQM(oo)
           frmWBXJ.lblQM(oo).Caption = ""
           Load frmWBXJ.cmdQm(oo)
           frmWBXJ.cmdQm(oo).Caption = ""
           Load frmWBXJ.lblTm(oo)
           frmWBXJ.lblTm(oo).Caption = ""
           frmWBXJ.lblQM(oo).Caption = mod1.HTP.Fields("QLabel").Value
            If mod1.HTP.Fields("xf").Value = True Then
               frmWBXJ.cmdQm(oo).Caption = mod1.HTP.Fields("Qren").Value
               frmWBXJ.lblTm(oo).Caption = mod1.HTP.Fields("QRQ").Value
            End If
           frmWBXJ.cmdQm(oo).Tag = mod1.HTP.Fields("zid").Value
           frmWBXJ.lblQM(oo).Visible = True
           frmWBXJ.cmdQm(oo).Visible = True
           frmWBXJ.lblTm(oo).Visible = True
           frmWBXJ.lblQM(oo).Left = frmWBXJ.lblQM(oo - 1).Left + 1100
           frmWBXJ.cmdQm(oo).Left = frmWBXJ.cmdQm(oo - 1).Left + 1100
           frmWBXJ.lblTm(oo).Left = frmWBXJ.lblTm(oo - 1).Left + 1100
           mod1.HTP.MoveNext
        Next
     Else
        frmWBXJ.cmdQm(0).Visible = False
        frmWBXJ.lblQM(0).Visible = False
        frmWBXJ.lblTm(0).Visible = False
     End If
Else                '购销
'''''''''    For oo = 10 To 1 Step -1
'''''''''        Unload frmGXBj.cmdQm(oo)
'''''''''        Unload frmGXBj.lblQM(oo)
'''''''''        Unload frmGXBj.lblTm(oo)
'''''''''    Next
'''''''''    frmGXBj.cmdQm(0).Caption = ""
'''''''''    frmGXBj.lblTm(0).Caption = ""
'''''''''      tt = "QMRZOpen(" & mod1.BTZ & ",'" & frmGXBj.lblBid.Caption & "')"
'''''''''      Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''      mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
'''''''''       If IsNull(mod1.HTP.RecordCount) = True Then
'''''''''            MsgBox ("连接超时,将退出!")
'''''''''            End
'''''''''        End If
'''''''''      If mod1.HTP.RecordCount > 0 Then
'''''''''         mod1.HTP.MoveFirst
'''''''''         frmGXBj.cmdQm(0).Visible = True
'''''''''         frmGXBj.lblQM(0).Visible = True
'''''''''         frmGXBj.lblTm(0).Visible = True
'''''''''         frmGXBj.lblQM(0).Caption = mod1.HTP.Fields("QLabel").Value
'''''''''         If mod1.HTP.Fields("xf").Value = True Then
'''''''''            frmGXBj.cmdQm(0).Caption = mod1.HTP.Fields("Qren").Value
'''''''''            frmGXBj.lblTm(0).Caption = mod1.HTP.Fields("QRQ").Value
'''''''''         End If
'''''''''         frmGXBj.cmdQm(0).Tag = mod1.HTP.Fields("zid").Value
'''''''''         mod1.HTP.MoveNext
'''''''''         For oo = 1 To mod1.HTP.RecordCount - 1
'''''''''           Load frmGXBj.lblQM(oo)
'''''''''           frmGXBj.lblQM(oo).Caption = ""
'''''''''           Load frmGXBj.cmdQm(oo)
'''''''''           frmGXBj.cmdQm(oo).Caption = ""
'''''''''           Load frmGXBj.lblTm(oo)
'''''''''           frmGXBj.lblTm(oo).Caption = ""
'''''''''            frmGXBj.lblQM(oo).Caption = mod1.HTP.Fields("QLabel").Value
'''''''''            If mod1.HTP.Fields("xf").Value = True Then
'''''''''               frmGXBj.cmdQm(oo).Caption = mod1.HTP.Fields("Qren").Value
'''''''''               frmGXBj.lblTm(oo).Caption = mod1.HTP.Fields("QRQ").Value
'''''''''            End If
'''''''''           frmGXBj.cmdQm(oo).Tag = mod1.HTP.Fields("zid").Value
'''''''''           frmGXBj.lblQM(oo).Visible = True
'''''''''           frmGXBj.cmdQm(oo).Visible = True
'''''''''           frmGXBj.lblTm(oo).Visible = True
'''''''''           frmGXBj.lblQM(oo).Left = frmGXBj.lblQM(oo - 1).Left + 1100
'''''''''           frmGXBj.cmdQm(oo).Left = frmGXBj.cmdQm(oo - 1).Left + 1100
'''''''''           frmGXBj.lblTm(oo).Left = frmGXBj.lblTm(oo - 1).Left + 1100
'''''''''           mod1.HTP.MoveNext
'''''''''        Next
'''''''''     Else
'''''''''        frmGXBj.cmdQm(0).Visible = False
'''''''''        frmGXBj.lblQM(0).Visible = False
'''''''''        frmGXBj.lblTm(0).Visible = False
'''''''''     End If


End If
End Sub
Public Sub OpenBJAN(LX As Boolean)
Dim tt As String
Dim oo As Integer
On Error Resume Next
If LX = True Then   '维保
    For oo = 10 To 1 Step -1
        Unload frmWbxjB.cmdQm(oo)
        Unload frmWbxjB.lblQM(oo)
        Unload frmWbxjB.lblTm(oo)
    Next
    frmWbxjB.cmdQm(0).Caption = ""
    frmWbxjB.lblTm(0).Caption = ""
      tt = "qmrzOpen(" & mod1.BTZ & ",'" & frmWbxjB.lblBaoId.Caption & "')"
      Set mod1.HTP = CreateObject("adodb.recordset")
      mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
      If mod1.HTP.RecordCount > 0 Then
         mod1.HTP.MoveFirst
         frmWbxjB.cmdQm(0).Visible = True
         frmWbxjB.lblQM(0).Visible = True
         frmWbxjB.lblTm(0).Visible = True
         frmWbxjB.lblQM(0).Caption = mod1.HTP.Fields("QLabel").Value
                      frmWbxjB.cmdQm(0).Tag = mod1.HTP.Fields("zid").Value
         If mod1.HTP.Fields("xf").Value = True Then
            frmWbxjB.cmdQm(0).Caption = mod1.HTP.Fields("Qren").Value
            frmWbxjB.lblTm(0).Caption = mod1.HTP.Fields("QRQ").Value

         End If

         mod1.HTP.MoveNext
         For oo = 1 To mod1.HTP.RecordCount - 1
           Load frmWbxjB.lblQM(oo)
           frmWbxjB.lblQM(oo).Caption = ""
           Load frmWbxjB.cmdQm(oo)
           frmWbxjB.cmdQm(oo).Caption = ""
           Load frmWbxjB.lblTm(oo)
           frmWbxjB.lblTm(oo).Caption = ""
           frmWbxjB.lblQM(oo).Caption = mod1.HTP.Fields("QLabel").Value
                           frmWbxjB.cmdQm(oo).Tag = mod1.HTP.Fields("zid").Value
           If mod1.HTP.Fields("xf").Value = True Then
                frmWbxjB.cmdQm(oo).Caption = mod1.HTP.Fields("Qren").Value
                frmWbxjB.lblTm(oo).Caption = mod1.HTP.Fields("QRQ").Value

           End If
           frmWbxjB.lblQM(oo).Visible = True
           frmWbxjB.cmdQm(oo).Visible = True
           frmWbxjB.lblTm(oo).Visible = True
           frmWbxjB.lblQM(oo).Left = frmWbxjB.lblQM(oo - 1).Left + 1100
           frmWbxjB.cmdQm(oo).Left = frmWbxjB.cmdQm(oo - 1).Left + 1100
           frmWbxjB.lblTm(oo).Left = frmWbxjB.lblTm(oo - 1).Left + 1100
           mod1.HTP.MoveNext
        Next
     Else
        frmWbxjB.cmdQm(0).Visible = False
        frmWbxjB.lblQM(0).Visible = False
        frmWbxjB.lblTm(0).Visible = False
     End If
ElseIf LX = False Then '购销
    For oo = 10 To 1 Step -1
        Unload frmGxbjB.cmdQm(oo)
        Unload frmGxbjB.lblQM(oo)
        Unload frmGxbjB.lblTm(oo)
    Next
    frmGxbjB.cmdQm(0).Caption = ""
    frmGxbjB.lblTm(0).Caption = ""
      tt = "qmrzOpen(" & mod1.BTZ & ",'" & frmGxbjB.lblBaoId.Caption & "')"
      Set mod1.HTP = CreateObject("adodb.recordset")
      mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
      If mod1.HTP.RecordCount > 0 Then
         mod1.HTP.MoveFirst
         frmGxbjB.cmdQm(0).Visible = True
         frmGxbjB.lblQM(0).Visible = True
         frmGxbjB.lblTm(0).Visible = True
        frmGxbjB.lblQM(0).Caption = mod1.HTP.Fields("QLabel").Value
                 frmGxbjB.cmdQm(0).Tag = mod1.HTP.Fields("zid").Value
        If mod1.HTP.Fields("xf").Value = True Then

         frmGxbjB.cmdQm(0).Caption = mod1.HTP.Fields("Qren").Value
         frmGxbjB.lblTm(0).Caption = mod1.HTP.Fields("QRQ").Value

        End If

         mod1.HTP.MoveNext
         For oo = 1 To mod1.HTP.RecordCount - 1
           Load frmGxbjB.lblQM(oo)
           frmGxbjB.lblQM(oo).Caption = ""
           Load frmGxbjB.cmdQm(oo)
           frmGxbjB.cmdQm(oo).Caption = ""
           Load frmGxbjB.lblTm(oo)
           frmGxbjB.lblTm(oo).Caption = ""
            frmGxbjB.lblQM(oo).Caption = mod1.HTP.Fields("QLabel").Value
                            frmGxbjB.cmdQm(oo).Tag = mod1.HTP.Fields("zid").Value
           If mod1.HTP.Fields("xf").Value = True Then

                frmGxbjB.cmdQm(oo).Caption = mod1.HTP.Fields("Qren").Value
                frmGxbjB.lblTm(oo).Caption = mod1.HTP.Fields("QRQ").Value

           End If

           frmGxbjB.lblQM(oo).Visible = True
           frmGxbjB.cmdQm(oo).Visible = True
           frmGxbjB.lblTm(oo).Visible = True
           frmGxbjB.lblQM(oo).Left = frmGxbjB.lblQM(oo - 1).Left + 1100
           frmGxbjB.cmdQm(oo).Left = frmGxbjB.cmdQm(oo - 1).Left + 1100
           frmGxbjB.lblTm(oo).Left = frmGxbjB.lblTm(oo - 1).Left + 1100
           mod1.HTP.MoveNext
        Next
     Else
        frmGxbjB.cmdQm(0).Visible = False
        frmGxbjB.lblQM(0).Visible = False
        frmGxbjB.lblTm(0).Visible = False
     End If
End If
End Sub



Public Sub BJDGDBound(Bid As Long) '在事务列表的状态下"直接"打开维保维修询价单的购销表(因为一般的打开方式为先打开维保,再打开购销)

Dim tt As String
Dim LX As Boolean
On Error Resume Next

        tt = "select * from xunJiaD where bid=" & Bid
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        frmGXBj.lblZl.Caption = mod1.HTP.Fields("zl").Value
        frmGXBj.comXmmc.Text = mod1.HTP.Fields("xmmc").Value
        frmGXBj.comXmmc.Tag = mod1.HTP.Fields("xid").Value
        frmGXBj.lblBid.Caption = mod1.HTP.Fields("bid").Value
        frmGXBj.lblOid.Caption = mod1.HTP.Fields("oid").Value
        frmGXBj.lblLc.Caption = mod1.HTP.Fields("lc").Value
        frmGXBj.lblLcRen.Caption = mod1.HTP.Fields("lcren").Value
        frmGXBj.lblLcUid.Caption = mod1.HTP.Fields("lcuid").Value
        frmGXBj.lblFwid.Caption = mod1.HTP.Fields("fwid").Value
        frmGXBj.lblNlb.Caption = mod1.HTP.Fields("nlb").Value
        frmGXBj.lblLcou.Caption = mod1.HTP.Fields("lcou").Value
        frmGXBj.lblBaoId.Caption = mod1.HTP.Fields("baoid").Value
        frmGXBj.lblBh.Caption = mod1.HTP.Fields("bianhao").Value
        frmGXBj.lblPwf.Caption = mod1.HTP.Fields("pwf").Value
        frmGXBj.txtHg.Text = mod1.HTP.Fields("hg").Value
        frmGXBj.txtYhg.Text = mod1.HTP.Fields("yhg").Value
        frmGXBj.lblYwy.Caption = mod1.HTP.Fields("ywy").Value
        frmGXBj.lblUid.Caption = mod1.HTP.Fields("uid").Value
        frmGXBj.lblWbid.Caption = mod1.HTP.Fields("wbid").Value
        If frmGXBj.lblZl.Caption = "大修" Then
            frmGXBj.cmdCT.Caption = mod1.HTP.Fields("CC").Value
            frmGXBj.lblCT.Caption = mod1.HTP.Fields("ctime").Value
            frmGXBj.frmCT.Visible = True
            frmGXBj.CTF = False
        End If
        
    If mod1.HTP.Fields("chf").Value = True And frmGXBj.lblLc.Caption > 2 Then
        frmGXBj.lblZ.Visible = True
        frmGXBj.lblZT.Visible = True
        frmGXBj.lblZT.Caption = mod1.HTP.Fields("chdate").Value

    End If
        frmGXBj.lblCfwid.Caption = mod1.HTP.Fields("cfwid").Value
        tt = "select qy,bm from RenYuan where userName='" & frmGXBj.lblYwy.Caption & "' and userid='" & frmGXBj.lblUid.Caption & "'"
        mod1.HTT.Close
        mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        frmGXBj.lblBM.Caption = mod1.HTT.Fields("bm").Value
        frmGXBj.lblQy.Caption = mod1.HTT.Fields("qy").Value
        
        
        tt = "select * from xunJIamxView where bid=" & Val(frmGXBj.lblBid.Caption)
        frmGXBj.adoGx.Close
        frmGXBj.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        If frmGXBj.lblZl.Caption = "大修" Then '检验是否要显示产品采购签字按钮
            frmGXBj.adoGx.MoveFirst
            Do While Not frmGXBj.adoGx.EOF
                If frmGXBj.adoGx.Fields("品种").Value = "产品" Then
                    frmGXBj.CTF = True
                    Exit Do
                End If
                frmGXBj.adoGx.MoveNext
            Loop
        End If
        If frmGXBj.CTF = False Then '如果不包含产品,则不显示产品签字按钮.
            frmGXBj.frmCT.Visible = False
        End If
        Set frmGXBj.dtgMa.DataSource = frmGXBj.adoGx
        If frmGXBj.adoGx.RecordCount > 1 Then
            frmGXBj.dtgMa.FixedRows = 0
            frmGXBj.dtgMa.MergeCol(1) = True
            frmGXBj.dtgMa.MergeCol(2) = True
            frmGXBj.dtgMa.MergeCol(10) = True
            frmGXBj.dtgMa.MergeCol(14) = True
            frmGXBj.dtgMa.MergeCells = 3
            frmGXBj.dtgMa.FixedRows = 1
        End If
        frmGXBj.cmdSave.Enabled = False
        frmGXBj.cmdMod.Enabled = True
        


        frmGXBj.cmdBjd.Visible = False
        frmGXBj.cmdMod.Enabled = True
        frmGXBj.cmdSave.Enabled = False
        Call modBJD.OpenXJAN(LX)

        If mod1.VLP = 2 Or mod1.VLP = 3 Then
            frmGXBj.cmdWb.Visible = False
        Else
            frmGXBj.cmdWb.Visible = True
        End If
'        If mod1.DName = frmGXBj.lblYwy.Caption Then
'            frmGXBj.cmdCong.Visible = True
'        Else
'            frmGXBj.cmdCong.Visible = False
'        End If


End Sub

Public Sub BJDWDBound(Bid As Long)  '在事务列表框打开编保合同询价的采购询价后,再打开相应的人工询价.
Dim tt As String
On Error Resume Next
tt = "select top 1 * from XunJiaD where bid=" & Bid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText

        frmWBXJ.cmdCG.Visible = True
        frmWBXJ.lblZl.Caption = mod1.HTP.Fields("Zl").Value
        frmWBXJ.comXmmc.Text = mod1.HTP.Fields("xmmc").Value
        frmWBXJ.comXmmc.Tag = mod1.HTP.Fields("xid").Value
        frmWBXJ.lblBid.Caption = mod1.HTP.Fields("bid").Value
        frmWBXJ.lblBh.Caption = mod1.HTP.Fields("bianhao").Value
        frmWBXJ.comZu.Text = mod1.HTP.Fields("zh").Value
        frmWBXJ.txtZu.Text = mod1.HTP.Fields("zName").Value
        frmWBXJ.comPb.Text = mod1.HTP.Fields("jzpb").Value
        frmWBXJ.comXh.Text = mod1.HTP.Fields("jzxh").Value
        frmWBXJ.txtSL.Text = mod1.HTP.Fields("sL").Value
        frmWBXJ.lblOid.Caption = mod1.HTP.Fields("oid").Value
        frmWBXJ.txtZT.Text = mod1.HTP.Fields("ZTime").Value
        frmWBXJ.txtClf.Text = mod1.HTP.Fields("clf").Value
        frmWBXJ.txtHg.Text = mod1.HTP.Fields("HG").Value
        frmWBXJ.txtYhg.Text = mod1.HTP.Fields("yhg").Value
        frmWBXJ.chkBa.Value = Abs(CInt(mod1.HTP.Fields("ta").Value))
        frmWBXJ.chkBb.Value = Abs(CInt(mod1.HTP.Fields("tb").Value))
        frmWBXJ.chkBc.Value = Abs(CInt(mod1.HTP.Fields("tc").Value))
        frmWBXJ.lblYwy.Caption = mod1.HTP.Fields("ywy").Value
        frmWBXJ.lblUid.Caption = mod1.HTP.Fields("uid").Value
        tt = "select qy,bm from RenYuan where userName='" & frmWBXJ.lblYwy.Caption & "' and userid='" & frmWBXJ.lblUid.Caption & "'"
        mod1.HTT.Close
        mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        frmWBXJ.lblBM.Caption = mod1.HTT.Fields("bm").Value
        frmWBXJ.lblQy.Caption = mod1.HTT.Fields("qy").Value
        frmWBXJ.lblBaoId.Caption = mod1.HTP.Fields("baoid").Value
        frmWBXJ.txtWc.Text = mod1.HTP.Fields("wc").Value
        frmWBXJ.txtXc.Text = mod1.HTP.Fields("Xc").Value
        frmWBXJ.txtMon.Text = mod1.HTP.Fields("mon").Value
        frmWBXJ.txtDxnr.Text = mod1.HTP.Fields("dxnr").Value
        frmWBXJ.lblCgid.Caption = mod1.HTP.Fields("cgid").Value
        frmWBXJ.lblPwf.Caption = mod1.HTP.Fields("pwf").Value
        frmWBXJ.lblLc.Caption = mod1.HTP.Fields("Lc").Value
        frmWBXJ.lblLcRen.Caption = mod1.HTP.Fields("LcRen").Value
        frmWBXJ.lblLcUid.Caption = mod1.HTP.Fields("LcUid").Value
        frmWBXJ.lblFwid.Caption = mod1.HTP.Fields("Fwid").Value
        frmWBXJ.lblNlb.Caption = mod1.HTP.Fields("Nlb").Value
        frmWBXJ.lblLcou.Caption = mod1.HTP.Fields("Lcou").Value
        If mod1.HTP.Fields("zl").Value = "维保" Then
            '年保表
            tt = "select * from xunJIaWbView where wbx='年保' and bid=" & Val(frmWBXJ.lblBid.Caption)
            frmWBXJ.adoWb.Close
            frmWBXJ.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            Set frmWBXJ.dtgWb.DataSource = frmWBXJ.adoWb
            frmWBXJ.dtgWb.FixedRows = 0
            frmWBXJ.dtgWb.MergeCol(1) = True
            frmWBXJ.dtgWb.MergeCol(2) = True
            frmWBXJ.dtgWb.MergeCol(3) = True
            frmWBXJ.dtgWb.MergeCells = 3
            frmWBXJ.dtgWb.FixedRows = 1
            '例检表
            tt = "select * from xunJIaWbView where wbx='例检' and bid=" & Val(frmWBXJ.lblBid.Caption)
            frmWBXJ.adoLj.Close
            frmWBXJ.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            Set frmWBXJ.dtgLj.DataSource = frmWBXJ.adoLj
            frmWBXJ.dtgLj.FixedRows = 0
            frmWBXJ.dtgLj.MergeCol(1) = True
            frmWBXJ.dtgLj.MergeCol(2) = True
            frmWBXJ.dtgLj.MergeCol(3) = True
            frmWBXJ.dtgLj.MergeCells = 3
            frmWBXJ.dtgLj.FixedRows = 1
            frmWBXJ.frmDx.Visible = False
            frmWBXJ.frmNb.Visible = True
            frmWBXJ.frmTime.Visible = True
            If frmWBXJ.lblLc = 1 Then   '业务员未审核前,还能修改维保条款,以后就不能

                frmWBXJ.cmdD.Visible = True
                frmWBXJ.cmdJi.Visible = True
            Else

                frmWBXJ.cmdD.Visible = False
                frmWBXJ.cmdJi.Visible = False

            End If
            frmWBXJ.tabGc.TabVisible(2) = False
            frmWBXJ.tabGc.TabVisible(0) = True
            frmWBXJ.tabGc.TabVisible(1) = True
            frmWBXJ.tabGc.Tab = 0
    
         Else '大修

            frmWBXJ.frmDx.Visible = True
            frmWBXJ.frmNb.Visible = False
            frmWBXJ.frmTime.Visible = False

            frmWBXJ.cmdD.Visible = False
            frmWBXJ.cmdJi.Visible = False
            frmWBXJ.tabGc.TabVisible(2) = True
            frmWBXJ.tabGc.TabVisible(0) = False
            frmWBXJ.tabGc.TabVisible(1) = False
            frmWBXJ.tabGc.Tab = 2
         End If
        frmWBXJ.cmdMod.Enabled = True
        frmWBXJ.cmdSave.Enabled = False


        If Val(frmWBXJ.lblLc.Caption) > 1 Then
            frmWBXJ.txtYhg.Locked = False
        Else
            frmWBXJ.txtYhg.Locked = True
        End If
        Call modBJD.OpenXJAN(True)
End Sub

Public Sub wbxjLocked()
'Set frmWBXJ.dtgWb.DataSource = Nothing
'Set frmWBXJ.dtgLj.DataSource = Nothing



frmWBXJ.comXmmc.Locked = True
frmWBXJ.comZu.Locked = True
frmWBXJ.txtZu.Locked = True
frmWBXJ.comPb.Locked = True
frmWBXJ.comXh.Locked = True
frmWBXJ.txtSL.Locked = True
frmWBXJ.txtZT.Locked = True
frmWBXJ.txtHg.Locked = True
frmWBXJ.txtYhg.Locked = True
frmWBXJ.chkBa.Enabled = False
frmWBXJ.chkBb.Enabled = False
frmWBXJ.chkBc.Enabled = False
frmWBXJ.txtWc.Locked = True
frmWBXJ.txtXc.Locked = True
frmWBXJ.txtClf.Locked = True
frmWBXJ.txtMon.Locked = True
frmWBXJ.txtDxnr.Locked = True
frmWBXJ.cmdJadd.Enabled = False
frmWBXJ.cmdJdel.Enabled = False
frmWBXJ.cmdJgx.Enabled = False
frmWBXJ.txtF1.Locked = True
frmWBXJ.txtF2.Locked = True
frmWBXJ.txtF3.Locked = True
frmWBXJ.txtF4.Locked = True
frmWBXJ.txtBz.Locked = True
frmWBXJ.txtFbje.Locked = True
frmWBXJ.txtFbnr.Locked = True
frmWBXJ.txt1.Locked = True
End Sub
Public Sub wbxjUnLocked()
'Set frmWBXJ.dtgWb.DataSource = Nothing
'Set frmWBXJ.dtgLj.DataSource = Nothing
frmWBXJ.comXmmc.Locked = False
frmWBXJ.comZu.Locked = False
frmWBXJ.txtZu.Locked = False
frmWBXJ.comPb.Locked = False
frmWBXJ.comXh.Locked = False
frmWBXJ.txtSL.Locked = False
frmWBXJ.txtZT.Locked = False
frmWBXJ.txtHg.Locked = False
frmWBXJ.txtYhg.Locked = False
frmWBXJ.chkBa.Enabled = True
frmWBXJ.chkBb.Enabled = True
frmWBXJ.chkBc.Enabled = True
frmWBXJ.txtWc.Locked = False
frmWBXJ.txtXc.Locked = False
frmWBXJ.txtMon.Locked = False
frmWBXJ.txtDxnr.Locked = False
frmWBXJ.txtClf.Locked = False
frmWBXJ.txtBz.Locked = False
frmWBXJ.txtFbje.Locked = False
frmWBXJ.txtFbnr.Locked = False

End Sub

Public Sub gxbjLocked()
On Error Resume Next
frmGXBj.comLx.Locked = True
frmGXBj.comXmmc.Locked = True
frmGXBj.comJzpb.Locked = True
frmGXBj.comJzXh.Locked = True
frmGXBj.txtYxh.Locked = True
frmGXBj.txtCbh.Locked = True
frmGXBj.txtCd.Locked = True
frmGXBj.txtLjbh.Locked = True
frmGXBj.txtLjmc.Locked = True
frmGXBj.txtXlh.Locked = True
frmGXBj.txtSL.Locked = True
frmGXBj.txtDj.Locked = True
frmGXBj.txtBrq.Locked = True
frmGXBj.txtMj.Locked = True

frmGXBj.txtHg.Locked = True
frmGXBj.txtYhg.Locked = True

frmGXBj.cmdQing.Enabled = False
frmGXBj.cmdAdd.Enabled = False
frmGXBj.cmdDel.Enabled = False
frmGXBj.cmdJgx.Enabled = False
frmGXBj.cmdGx.Enabled = False
frmGXBj.txtBz.Locked = True
frmGXBj.txtYf.Locked = True
frmGXBj.txtADR.Locked = True
frmGXBj.cmdGsave.Enabled = False
End Sub

Public Sub gxbjUnLocked()
On Error Resume Next

frmGXBj.comXmmc.Locked = False
frmGXBj.comJzpb.Locked = False
frmGXBj.comJzXh.Locked = False
frmGXBj.txtYxh.Locked = False
frmGXBj.txtCbh.Locked = False
frmGXBj.txtCd.Locked = False
frmGXBj.txtLjbh.Locked = False
frmGXBj.txtLjmc.Locked = False
frmGXBj.txtXlh.Locked = False
frmGXBj.txtSL.Locked = False
frmGXBj.txtDj.Locked = False
frmGXBj.txtBrq.Locked = False
frmGXBj.txtMj.Locked = False

'frmGXBj.txtHg.Locked = False
'frmGXBj.txtYhg.Locked = False

frmGXBj.cmdQing.Enabled = True
frmGXBj.cmdAdd.Enabled = True
frmGXBj.cmdDel.Enabled = True
frmGXBj.cmdJgx.Enabled = True
frmGXBj.cmdGx.Enabled = True
frmGXBj.txtBz.Locked = False
End Sub

Public Sub gxbjbLocked()

frmGxbjB.comXmmc.Locked = True

frmGxbjB.comKhmc.Locked = True

frmGxbjB.txtDj.Locked = True
frmGxbjB.txtSL.Locked = True
frmGxbjB.txtHg.Locked = True
frmGxbjB.txtYhg.Locked = True
frmGxbjB.txtYJ.Locked = True
frmGxbjB.txtTcBe.Locked = True
frmGxbjB.optLa.Enabled = False
frmGxbjB.optLb.Enabled = False
frmGxbjB.optLc.Enabled = False
frmGxbjB.txtFbje.Locked = True

frmGxbjB.txtXm1.Locked = True
frmGxbjB.txtXm2.Locked = True
frmGxbjB.txtClcb.Locked = True
frmGxbjB.txtYJ.Locked = True
frmGxbjB.txtYf.Locked = True
frmGxbjB.txtCb.Locked = True
frmGxbjB.cmdGx.Enabled = False
frmGxbjB.txtBz.Locked = True
End Sub

Public Sub wbxjbLocked()
frmWbxjB.comXmmc.Locked = True


frmWbxjB.comZu.Locked = True
frmWbxjB.txtZu.Locked = True
frmWbxjB.comPb.Locked = True
frmWbxjB.comXh.Locked = True
frmWbxjB.txtSL.Locked = True

frmWbxjB.txtZT.Locked = True
frmWbxjB.txtHg.Locked = True
frmWbxjB.txtYhg.Locked = True
frmWbxjB.chkBa.Enabled = False
frmWbxjB.chkBb.Enabled = False
frmWbxjB.chkBc.Enabled = False
frmWbxjB.txtTl.Locked = True
frmWbxjB.cmdLeft.Enabled = False
frmWbxjB.cmdRight.Enabled = False
frmWbxjB.txtFbje.Locked = True

frmWbxjB.txtRgf.Locked = True
frmWbxjB.txtClf.Locked = True
frmWbxjB.txtClcb.Locked = True
frmWbxjB.txtYJ.Locked = True
frmWbxjB.txtMon.Locked = True
frmWbxjB.txtWc.Locked = True
frmWbxjB.txtXc.Locked = True

frmWbxjB.txtXm1.Locked = True
frmWbxjB.txtXm2.Locked = True
frmWbxjB.txtYf.Locked = True
frmWbxjB.txtYJ.Locked = True
frmWbxjB.txtTcBe.Locked = True
frmWbxjB.dt3.Enabled = False
frmWbxjB.dt4.Enabled = False
frmWbxjB.txtBz.Locked = True
End Sub
