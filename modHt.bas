Attribute VB_Name = "modHt"
Public hTchar As String 'Sql����
Public adoRGF As Object

Public Sub addHt() '��ͬ�����ύ



End Sub

Public Sub htQing() '��ͬ�������
form2Htp.lblKhdh.Caption = ""
form2Htp.txtKhmc.Text = "" '�ͻ�����
form2Htp.txtXmmc.Text = "" '��Ŀ����
form2Htp.txtYwy.Text = "" 'ҵ��Ա
form2Htp.txtHtbh.Text = "" '��ͬ���
form2Htp.dt1.Value = mod1.HMDa
'form2Htp.comQy.Text = "" '����
form2Htp.txtHtdate.Text = ""  '��ͬ����
'form2Htp.txtDdbh.Text = "" '�������
'form2Htp.txtDddate.Text = "" '��������
form2Htp.lblHtxz.Caption = "" '��ͬ����
form2Htp.optA(0).Value = False
form2Htp.optA(1).Value = False
form2Htp.optA(5).Value = False
'form2Htp.txtHtqy.Text = "" '��ͬ������ʼ��
'form2Htp.txtHtqy1.Text = "" '��ͬ���޽�����
'form2Htp.comJzpb.Text = " " '����Ʒ��
'form2Htp.txtJzxh.Text = "" '�����ͺ�
'form2Htp.txtJzcount.Text = "" '��������
'form2Htp.txtYjxh.Text = "" 'ѹ�����ͺ�
form2Htp.txtTian.Text = "" '������
form2Htp.txtJhqk.Text = "" '�������
form2Htp.txtMOn.Text = "" '������
form2Htp.txtHtze.Text = "" '��ͬ�ܶ�
form2Htp.txtClf.Text = "" '���Ϸ�
form2Htp.txtRgf.Text = "" '�˹���
form2Htp.txtCbze1.Text = "" '�ɱ��ܶ�
form2Htp.txtCbze2.Text = "" 'ʵ�ʳɱ��ܶ�
'form2Htp.txtCbze3.Text = "" '��ͬ�ɱ��ܶ�
form2Htp.txtClcb1.Text = "" '���ϳɱ�
form2Htp.txtClcb2.Text = "" 'ʵ�ʲ��ϳɱ�
'form2Htp.txtClcb3.Text = "" '��ͬ���ϳɱ�
form2Htp.txtFbje1.Text = "" '�ְ����
form2Htp.txtFbje2.Text = "" 'ʵ�ʷְ����

'form2Htp.txtFbje3.Text = "" '��ͬ�ְ����
form2Htp.txtYf1.Text = "" '�˷�
form2Htp.txtYf2.Text = "" 'ʵ���˷�
'form2Htp.txtYf3.Text = "" '��ͬ�˷�
form2Htp.txtQt1.Text = "" '����
form2Htp.txtQt2.Text = "" 'ʵ������
'form2Htp.txtQt3.Text = "" '��ͬ����
form2Htp.txtYj1.Text = "" 'Ӷ��
form2Htp.txtYj2.Text = "" 'ʵ��Ӷ��
'form2Htp.txtYj3.Text = "" '��ͬӶ��
form2Htp.txtLr1.Text = "" '��Ŀ����
form2Htp.txtLr2.Text = "" 'ʵ����Ŀ����
'form2Htp.txtLr3.Text = "" '��ͬ��Ŀ����
form2Htp.txtTc1.Text = "" '���
form2Htp.txtTc2.Text = "" 'ʵ�����
'form2Htp.txtTc3.Text = "" '��ͬ���
form2Htp.txtJlr1.Text = ""
form2Htp.txtJlr2.Text = ""
form2Htp.txtZXF1.Text = "" 'װж��
form2Htp.txtZxF2.Text = "" 'ʵ��װж��

form2Htp.txtCBze3.Text = ""
form2Htp.txtClcb3.Text = ""
form2Htp.txtQT3.Text = ""
form2Htp.txtZXF3.Text = ""
form2Htp.txtYf3.Text = ""
form2Htp.txtFbje3.Text = ""
form2Htp.txtTcBe.Text = 6 '��ɱ���
form2Htp.txtTcBe.Visible = False
form2Htp.lblTcBe.Visible = False
form2Htp.UpDa.Visible = False
form2Htp.txtTc1.Visible = True
form2Htp.txtTc2.Visible = True

'��Ʊ����
form2Htp.optLa.Value = False
form2Htp.optLb.Value = False
form2Htp.optLc.Value = False
form2Htp.optLD.Value = False
form2Htp.optLE.Value = False
'form2Htp.frmKP.Visible = False
form2Htp.chkDzf.Value = False

form2Htp.chkA.Caption = "" 'ҵ��Աǩ��
form2Htp.chkA.Tag = "" '����ҵ��Ա
form2Htp.chkB.Caption = "" '���۾���ǩ��
form2Htp.chkB.Tag = "" '�������۾���
form2Htp.chkC.Caption = "" '������ǩ��
form2Htp.chkD.Caption = "" '�ܾ���ǩ��
form2Htp.chkE.Caption = "" '����֧��ǩ��
form2Htp.lblYw.Caption = ""
form2Htp.lblJZ.Caption = ""
form2Htp.lblJl.Caption = ""
form2Htp.lblYZ.Caption = ""
form2Htp.lblZJ.Caption = ""

'����ֶ�
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
form2Htp.txtFkBz.Text = "" '����������ע
form2Htp.lblHid.Caption = ""
form2Htp.txtTcRQ.Text = "���ȡ������"

form2Htp.optP.BackColor = &H8000000F
End Sub


Public Sub htBound() '��ͬ�����ֶΰ�(ֱ�Ӹ�ֵ�������ð󶨷���)
On Error Resume Next
Call mod1.zhuDa(1, mod1.HTP.Fields("htbh").Value)
form2Htp.txtKhmc.Text = mod1.HTP.Fields(0).Value '�ͻ�����
form2Htp.lblKhdh.Caption = mod1.HTP.Fields("khdh").Value '�ͻ�����
form2Htp.txtXmmc.Text = mod1.HTP.Fields("xmmc").Value '��Ŀ����
form2Htp.txtYwy.Text = mod1.HTP.Fields("YwY").Value 'ҵ��Ա
form2Htp.txtHtbh.Text = mod1.HTP.Fields(2).Value '��ͬ���

form2Htp.txtHtdate.Text = Format(mod1.HTP.Fields(3).Value, "Long Date") '��ͬ����
form2Htp.dt1.Value = mod1.HTP.Fields(3).Value
'form2Htp.txtDdbh.Text = mod1.HtP.Fields(4).Value '�������
'form2Htp.txtDddate.Text = Format(mod1.HtP.Fields(5).Value, "Long Date") '��������
'form2Htp.dt2.Value = mod1.HtP.Fields(5).Value

form2Htp.lblHtxz.Caption = mod1.HTP.Fields(6).Value '��ͬ����
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
Case "A. �������ͬ"
form2Htp.optA(0).Value = True
Case "B1.���̺�ͬ"
form2Htp.optA(1).Value = True
'Case "B2.�ְ���ͬ"
'form2Htp.optA(2).Value = True
Case "C. ά����ͬ"
form2Htp.optA(3).Value = True
Case "D. ά�޺�ͬ"
form2Htp.optA(4).Value = True
Case "E. ��Ʒ��ͬ"
form2Htp.optA(5).Value = True
End Select

form2Htp.comQy.Text = mod1.HTP.Fields("qy").Value '����

'form2Htp.txtHtqy.Text = Format(mod1.HtP.Fields(7).Value, "Long Date") '��ͬ������ʼ��
'form2Htp.dt3.Value = mod1.HtP.Fields(7).Value
'form2Htp.txtHtqy1.Text = Format(mod1.HtP.Fields(8).Value, "Long Date") '��ͬ���޽�����
'form2Htp.dt4.Value = mod1.HtP.Fields(8).Value
'form2Htp.comJzpb.Text = mod1.HtP.Fields(9).Value '����Ʒ��
'form2Htp.txtJzxh.Text = mod1.HtP.Fields(10).Value '�����ͺ�
'form2Htp.txtJzcount.Text = mod1.HtP.Fields(11).Value '��������
'form2Htp.txtYjxh.Text = mod1.HtP.Fields(12).Value 'ѹ�����ͺ�
form2Htp.txtTian.Text = mod1.HTP.Fields(13).Value '������
form2Htp.txtJhqk.Text = mod1.HTP.Fields(14).Value '�������
form2Htp.txtMOn.Text = mod1.HTP.Fields(15).Value '������
If mod1.HTP.Fields(16).Value = 0 Then
form2Htp.txtHtze.Text = ""
Else
form2Htp.txtHtze.Text = mod1.HTP.Fields(16).Value '��ͬ�ܶ�
End If

frmFuK.lblHtze.Caption = Round(mod1.HTP.Fields(16).Value, 2)

If mod1.HTP.Fields(17).Value = 0 Then
form2Htp.txtClf.Text = ""
Else
form2Htp.txtClf.Text = mod1.HTP.Fields(17).Value '���Ϸ�
End If

If mod1.HTP.Fields(18).Value = 0 Then
form2Htp.txtRgf.Text = ""
Else
form2Htp.txtRgf.Text = mod1.HTP.Fields(18).Value '�˹���
End If

If mod1.HTP.Fields(19).Value = 0 Then
form2Htp.txtCbze1.Text = ""
Else
form2Htp.txtCbze1.Text = mod1.HTP.Fields(19).Value '�ɱ��ܶ�
End If

If mod1.HTP.Fields(20).Value = 0 Then
form2Htp.txtCbze2.Text = ""
Else
form2Htp.txtCbze2.Text = mod1.HTP.Fields("cbze1").Value 'ʵ�ʳɱ��ܶ�
End If

If mod1.HTP.Fields(52).Value = 0 Then
'form2Htp.txtCbze3.Text = ""
Else
'form2Htp.txtCbze3.Text = mod1.HtP.Fields(52).Value '��ͬ�ɱ��ܶ�
End If

If mod1.HTP.Fields(21).Value = 0 Then
form2Htp.txtClcb1.Text = ""
Else
form2Htp.txtClcb1.Text = mod1.HTP.Fields(21).Value '���ϳɱ�
End If

If mod1.HTP.Fields(22).Value = 0 Then
form2Htp.txtClcb2.Text = ""
Else
form2Htp.txtClcb2.Text = mod1.HTP.Fields(22).Value 'ʵ�ʲ��ϳɱ�
End If

If mod1.HTP.Fields(53).Value = 0 Then
'form2Htp.txtClcb3.Text = ""
Else
'form2Htp.txtClcb3.Text = mod1.HtP.Fields(53).Value '��ͬ���ϳɱ�
End If

If mod1.HTP.Fields(23).Value = 0 Then
form2Htp.txtFbje1.Text = ""
Else
form2Htp.txtFbje1.Text = mod1.HTP.Fields(23).Value '�ְ����
End If

If mod1.HTP.Fields(24).Value = 0 Then
form2Htp.txtFbje2.Text = ""
Else
form2Htp.txtFbje2.Text = mod1.HTP.Fields(24).Value 'ʵ�ʷְ����
End If

If mod1.HTP.Fields(54).Value = 0 Then
'form2Htp.txtFbje3.Text = ""
Else
'form2Htp.txtFbje3.Text = mod1.HtP.Fields(54).Value '��ͬ�ְ����
End If


If mod1.HTP.Fields(25).Value = 0 Then
form2Htp.txtYf1.Text = ""
Else
form2Htp.txtYf1.Text = mod1.HTP.Fields(25).Value '�˷�
End If

If mod1.HTP.Fields(26).Value = 0 Then
form2Htp.txtYf2.Text = ""
Else
form2Htp.txtYf2.Text = mod1.HTP.Fields(26).Value 'ʵ���˷�
End If

If mod1.HTP.Fields(55).Value = 0 Then
'form2Htp.txtYf3.Text = ""
Else
'form2Htp.txtYf3.Text = mod1.HtP.Fields(55).Value '��ͬ�˷�
End If

If mod1.HTP.Fields(27).Value = 0 Then
form2Htp.txtQt1.Text = ""
Else
form2Htp.txtQt1.Text = mod1.HTP.Fields(27).Value '����
End If

If mod1.HTP.Fields(28).Value = 0 Then
form2Htp.txtQt2.Text = ""
Else
form2Htp.txtQt2.Text = mod1.HTP.Fields(28).Value 'ʵ������
End If

If mod1.HTP.Fields(56).Value = 0 Then
'form2Htp.txtQt3.Text = ""
Else
'form2Htp.txtQt3.Text = mod1.HtP.Fields(56).Value '��ͬ����
End If

If mod1.HTP.Fields(29).Value = 0 Then
form2Htp.txtYj1.Text = ""
Else
form2Htp.txtYj1.Text = mod1.HTP.Fields(29).Value 'Ӷ��
End If

If mod1.HTP.Fields(30).Value = 0 Then
form2Htp.txtYj2.Text = ""
Else
form2Htp.txtYj2.Text = mod1.HTP.Fields(30).Value 'ʵ��Ӷ��
End If

If mod1.HTP.Fields(57).Value = 0 Then
'form2Htp.txtYj3.Text = ""
Else
'form2Htp.txtYj3.Text = mod1.HtP.Fields(57).Value '��ͬӶ��
End If

If mod1.HTP.Fields(31).Value = 0 Then
form2Htp.txtLr1.Text = ""
Else
form2Htp.txtLr1.Text = mod1.HTP.Fields(31).Value '��Ŀ����
End If

If mod1.HTP.Fields(32).Value = 0 Then
form2Htp.txtLr2.Text = ""
Else
form2Htp.txtLr2.Text = mod1.HTP.Fields(32).Value 'ʵ����Ŀ����
End If

If mod1.HTP.Fields(58).Value = 0 Then
'form2Htp.txtLr3.Text = ""
Else
'form2Htp.txtLr3.Text = mod1.HtP.Fields(58).Value '��ͬ��Ŀ����
End If

form2Htp.txtJlr1.Text = mod1.HTP.Fields("jlr1").Value
form2Htp.txtJlr2.Text = mod1.HTP.Fields("jlr2").Value



form2Htp.txtCBze3.Text = mod1.HTP.Fields("cbze3").Value
form2Htp.txtClcb3.Text = mod1.HTP.Fields("clcb3").Value
form2Htp.txtQT3.Text = mod1.HTP.Fields("qt3").Value
form2Htp.txtZXF3.Text = mod1.HTP.Fields("zxf3").Value
form2Htp.txtYf3.Text = mod1.HTP.Fields("yf3").Value
form2Htp.txtFbje3.Text = mod1.HTP.Fields("fbje3").Value
form2Htp.txtTcBe.Text = mod1.HTP.Fields("tcbe").Value '��ɱ���

If mod1.HTP.Fields(33).Value = 0 Then
form2Htp.txtTc1.Text = ""
Else
form2Htp.txtTc1.Text = mod1.HTP.Fields(33).Value '���
End If

If mod1.HTP.Fields(34).Value = 0 Then
form2Htp.txtTc2.Text = ""
Else
form2Htp.txtTc2.Text = mod1.HTP.Fields(34).Value 'ʵ�����
End If

If mod1.HTP.Fields(59).Value = 0 Then
'form2Htp.txtTc3.Text = ""
Else
'form2Htp.txtTc3.Text = mod1.HtP.Fields(59).Value '��ͬ���
End If

form2Htp.txtZXF1.Text = mod1.HTP.Fields("rgF").Value 'Ԥ��װж��
form2Htp.txtZxF2.Text = mod1.HTP.Fields("rgF1").Value 'ʵ��װж��

'��Ʊ����
If mod1.HTP.Fields("fpLX").Value = "��ֵ��Ʊ" Then
form2Htp.optLa.Value = True
ElseIf mod1.HTP.Fields("fpLX").Value = "��ҵ��Ʊ" Then
form2Htp.optLb.Value = True
ElseIf mod1.HTP.Fields("fpLX").Value = "����Ʊ" Then
form2Htp.optLc.Value = True
ElseIf mod1.HTP.Fields("fpLX").Value = "����" Then
form2Htp.optLD.Value = True
ElseIf mod1.HTP.Fields("fpLX").Value = "����Ʊ" Then
form2Htp.optLE.Value = True
End If
'ĩ��Ʊ����
If mod1.HTP.Fields("dzF").Value = 1 Then
    form2Htp.chkDzf.Value = 1
ElseIf mod1.HTP.Fields("dzF").Value = 0 Then
    form2Htp.chkDzf.Value = 0
End If

If mod1.HTP.Fields(35).Value <> "" Then 'ҵ��Աǩ��
form2Htp.chkA.Caption = mod1.HTP.Fields(35).Value
form2Htp.chkA.Tag = mod1.HTP.Fields("xywy").Value
form2Htp.chkA.Value = 1
Else
form2Htp.chkA.Caption = ""
form2Htp.chkA.Value = 0
End If

If mod1.HTP.Fields("JzQz").Value <> "" Then '����֧��ǩ��
form2Htp.chkE.Caption = mod1.HTP.Fields("JzQz").Value
form2Htp.chkE.Value = 1
Else
form2Htp.chkE.Caption = ""
form2Htp.chkE.Value = 0
End If

If mod1.HTP.Fields(36).Value <> "" Then '���۾���ǩ��
form2Htp.chkB.Caption = mod1.HTP.Fields(36).Value
form2Htp.chkB.Tag = mod1.HTP.Fields("xjlq").Value
form2Htp.chkB.Value = 1
Else
form2Htp.chkB.Caption = ""
form2Htp.chkB.Value = 0
End If
If mod1.HTP.Fields(37).Value <> "" Then '������ǩ��
form2Htp.chkC.Caption = mod1.HTP.Fields(37).Value
form2Htp.chkC.Value = 1
Else
form2Htp.chkC.Caption = ""
form2Htp.chkC.Value = 0
End If
If mod1.HTP.Fields(38).Value <> "" Then '�ܾ���ǩ��
form2Htp.chkD.Caption = mod1.HTP.Fields(38).Value
form2Htp.chkD.Value = 1
Else
form2Htp.chkD.Caption = ""
form2Htp.chkD.Value = 0
End If

form2Htp.lblYw.Caption = mod1.HTP.Fields("ywDa").Value '����ǩ������
form2Htp.lblJZ.Caption = mod1.HTP.Fields("JzDa").Value '����֧��ǩ������
form2Htp.lblJl.Caption = mod1.HTP.Fields("JlDa").Value '���۾���ǩ������
form2Htp.lblYZ.Caption = mod1.HTP.Fields("YzDa").Value '������ǩ������
form2Htp.lblZJ.Caption = mod1.HTP.Fields("ZjDa").Value '�ܾ���ǩ������

'form2Htp.txtYwy.Text = form2Htp.chkA.Caption  'ҵ��Ա

'��ӵ��ֶ�
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

form2Htp.txtFkBz.Text = mod1.HTP.Fields("fkBz").Value '����������ע
form2Htp.comQy.Text = mod1.HTP.Fields("qy").Value '����
form2Htp.lblBM.Caption = mod1.HTP.Fields("bm").Value

If IsNull(mod1.HTP.Fields("TCRQ").Value) = False Then
    form2Htp.txtTcRQ.Text = mod1.HTP.Fields("TCRQ").Value '���ȡ������
End If

If mod1.HTP.Fields("jTf").Value = True Then
    form2Htp.cmdCount.Caption = "�ѽ���"
    form2Htp.cmdCount.Enabled = False
Else
    form2Htp.cmdCount.Caption = "����"
    form2Htp.cmdCount.Enabled = True
End If

'���Ϊ�ɺ�ͬ,������׶���Ϊ��ɫ
If mod1.HTP.Fields("XGG").Value = True Then
    form2Htp.optP.BackColor = &HC0FFFF
ElseIf mod1.HTP.Fields("XGG").Value = 0 Then
    form2Htp.optP.BackColor = &H8000000F
End If

form2Htp.optP.Enabled = False
form2Htp.optG.Enabled = False
form2Htp.optZ.Enabled = False
form2Htp.optW.Enabled = False

'��ִͬ�з�
If mod1.HTP.Fields(51).Value = 0 Then
    form2Htp.optP.Value = True
'    form2Htp.optP.Enabled = True
'    '�����ǩ����,����Ϊ�����ߴ�,����Խ��и���
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
''�����ͬ����û��ȫͨ�����߲���С�⣬�򡰺�ִͬ�з��ֶβ��ܱ༭
'If form2Htp.chkD.Caption <> "" And frmLogin.Combo1.Text = "��ӱ" Then
'form2Htp.frmZt.Enabled = True
'Else
'ĩ��Ʊ���ʷ�
If mod1.HTP.Fields("dzF").Value = True Then
    form2Htp.chkDzf.Value = 1
ElseIf mod1.HTP.Fields("dzF").Value = False Then
    form2Htp.chkDzf.Value = 0
End If
End Sub























Public Sub lianJ() '�ж��տ����ÿ����¼�Ƿ���Ӧ�ձ���һһ��Ӧ���������ɾ��

End Sub




Public Sub HtF() '�ж�htping,htping1,yiFk,htSale���е�htF,�Ƿ�һ�£��������ȫ����Ӧhtping��ֵ


End Sub























Public Sub qianKuan() '����Ƿ��
Dim ladate As Date
Dim cadate As Integer
Dim LT As String
On Error Resume Next
'������տ��ȷ����󸶿�����,�ٽ���Ƿ��ͳ��
If frmFuK.adoYf.Recordset.Fields(5).Value = True Then
    frmFuK.adoYf.Recordset.MoveLast
    ladate = frmFuK.adoYf.Recordset.Fields(0).Value
        Do While Not frmFuK.adoYf.Recordset.BOF
        cadate = DateDiff("D", frmFuK.adoYf.Recordset.Fields(3).Value, ladate)
        frmFuK.adoYf.Recordset.Update "laRq", ladate
        '����Ƿ�����
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
        
'�����δ�գ�����ݵ�ǰ������ȷ��Ƿ��
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

'�����ʽ��������Ƿ��
frmFuK.adoYf.Recordset.MoveFirst
Do While Not frmFuK.adoYf.Recordset.EOF
LT = "update llb1 set qianKuan1=1 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'" '����
If frmFuK.adoYf.Recordset.Fields(12).Value = 0 Then
LT = ""
LT = "update llb1 set qianKuan1=0 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
LT = ""
LT = "update llb1 set qianKuan2=1 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'" '����
If frmFuK.adoYf.Recordset.Fields(13).Value = 0 Then
LT = ""
LT = "update llb1 set qianKuan2=0 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
LT = ""
LT = "update llb1 set qianKuan3=1 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'" '3��
If frmFuK.adoYf.Recordset.Fields(14).Value = 0 Then
LT = ""
LT = "update llb1 set qianKuan3=0 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
LT = ""
LT = "update llb1 set qianKuan4=1 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'" '����
If frmFuK.adoYf.Recordset.Fields(15).Value = 0 Then
LT = ""
LT = "update llb1 set qianKuan4=0 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
LT = ""
LT = "update llb1 set qianKuan5=1 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'" '1��
If frmFuK.adoYf.Recordset.Fields(16).Value = 0 Then
LT = ""
LT = "update llb1 set qianKuan5=0 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
LT = ""
LT = "update llb1 set qianKuan6=1 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'" '2��
If frmFuK.adoYf.Recordset.Fields(17).Value = 0 Then
LT = ""
LT = "update llb1 set qianKuan6=0 where rq='" & frmFuK.adoYf.Recordset.Fields(0).Value & "'"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open LT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

frmFuK.adoYf.Recordset.MoveNext
Loop
End Sub

Public Sub tgYik(daDate As Variant)    'ͳ�Ƶ�ǰӦ��

End Sub


Public Sub gxAdd() '���ۺ�ͬ�������


End Sub
















































Public Sub gxQing() '���ۺ�ͬ�ֶ����
Dim oo As Integer
htgX.txtGF.Text = "" '����
htgX.comKhmc.Text = "" '�ͻ�����
If htgX.comKhmc.ListCount > 0 Then
    For oo = htgX.comKhmc.ListCount - 1 To 0 Step -1
        htgX.comKhmc.RemoveItem oo
    Next
End If
htgX.comKhmc.Locked = False
htgX.txtHtbh.Text = "" '��ͬ���
htgX.txtQyDD.Text = "" 'ǩԼ�ص�
htgX.txtXF.Text = "" '�跽
'htgX.DTPQdDate.Value = "" 'ǩ��ʱ��
htgX.txtHg.Text = "" '�ϼ�
htgX.lblDx.Caption = "" '�ϼƴ�д
htgX.txtT2.Text = "" '��������Ҫ������׼
htgX.txtZBQ.Text = "" '������
htgX.txtT3.Text = "" '�����������������������������
'htgX.txtT4.Text = "" '�ġ���(��)����ʽ
htgX.txtT5.Text = "" '�塢���䷽ʽ������վ���ۣ��ķ��ø���
htgX.txtT6.Text = "" '����������ļ��㷽��
htgX.txtT7.Text = "" '�ߡ���װ��׼����װ��Ĺ�Ӧ����պͷ��ø���
htgX.txtT8.Text = "" '�ˡ����շ�ʽ�������������
htgX.txtT9.Text = "" '�š������Ʒ�����������������Ӧ�취
htgX.txtT10.Text = "" 'ʮ�����㷽ʽ������
htgX.txtT11.Text = "" 'ʮһ�������ṩ������������ͬ�����飬��Ϊ����ͬ����
htgX.txtT12.Text = "" 'ʮ����ΥԼ����
htgX.txtT13.Text = "" 'ʮ���������ͬ���׵ķ�ʽ
htgX.txtT14.Text = "" 'ʮ�ġ�����Լ������
htgX.txtGdwMc.Text = "" '��λ����
htgX.txtGdwAdr.Text = "" '��λ��ַ
htgX.txtGfdBr.Text = "" '����������
htgX.txtGdw.Text = "" '�绰
htgX.txtGFX.Text = "" '����
htgX.txtGFH.Text = "" '��˰��
htgX.txtGkhYY.Text = "" '��������
htgX.txtGZH.Text = "" '�˺�
htgX.txtGyzBM.Text = "" '��������
htgX.txtGwiTo.Text = "" 'ί�д�����
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
'htgX.dtpYXQ.Value = "" '��Ч��

End Sub

























Public Sub gxBound() '���ۺ�ͬ�ֶΰ�
Dim tt As String
On Error Resume Next
tt = "Select * from gxHt where htBh='" & form2Htp.txtHtbh.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
htgX.txtGF.Text = mod1.HTP.Fields("GF").Value '����
htgX.txtHtbh.Text = mod1.HTP.Fields("htBh").Value '��ͬ���
htgX.txtQyDD.Text = mod1.HTP.Fields("qyDD").Value 'ǩԼ�ص�
htgX.txtXF.Text = mod1.HTP.Fields("XF").Value '�跽
htgX.comKhmc.Text = mod1.HTP.Fields("khmc").Value '�ͻ�����
htgX.DTPQdDate.Value = mod1.HTP.Fields("qdDate").Value 'ǩ��ʱ��
htgX.txtHg.Text = mod1.HTP.Fields("HG").Value '�ϼ�
htgX.lblDx.Caption = mod1.HTP.Fields("DHG").Value '�ϼƴ�д
htgX.txtT2.Text = mod1.HTP.Fields("T2").Value '��������Ҫ������׼
htgX.txtZBQ.Text = mod1.HTP.Fields("ZBQ").Value '������
htgX.txtT3.Text = mod1.HTP.Fields("T3").Value '�����������������������������
htgX.txtT4.Text = mod1.HTP.Fields("T4").Value '�ġ���(��)����ʽ
htgX.txtT5.Text = mod1.HTP.Fields("T5").Value '�塢���䷽ʽ������վ���ۣ��ķ��ø���
htgX.txtT6.Text = mod1.HTP.Fields("T6").Value '����������ļ��㷽��
htgX.txtT7.Text = mod1.HTP.Fields("T7").Value '�ߡ���װ��׼����װ��Ĺ�Ӧ����պͷ��ø���
htgX.txtT8.Text = mod1.HTP.Fields("T8").Value '�ˡ����շ�ʽ�������������
htgX.txtT9.Text = mod1.HTP.Fields("T9").Value '�š������Ʒ�����������������Ӧ�취
htgX.txtT10.Text = mod1.HTP.Fields("T10").Value 'ʮ�����㷽ʽ������
htgX.txtT11.Text = mod1.HTP.Fields("T11").Value 'ʮһ�������ṩ������������ͬ�����飬��Ϊ����ͬ����
htgX.txtT12.Text = mod1.HTP.Fields("T12").Value 'ʮ����ΥԼ����
htgX.txtT13.Text = mod1.HTP.Fields("T13").Value 'ʮ���������ͬ���׵ķ�ʽ
htgX.txtT14.Text = mod1.HTP.Fields("T14").Value 'ʮ�ġ�����Լ������
htgX.txtGdwMc.Text = mod1.HTP.Fields("GdwMc").Value  '��λ����
htgX.txtGdwAdr.Text = mod1.HTP.Fields("GdwAdr").Value '��λ��ַ
htgX.txtGfdBr.Text = mod1.HTP.Fields("GfdBr").Value '����������
htgX.txtGdw.Text = mod1.HTP.Fields("GdW").Value '�绰
htgX.txtGFX.Text = mod1.HTP.Fields("GFX").Value '����
htgX.txtGFH.Text = mod1.HTP.Fields("GFH").Value '��˰��
htgX.txtGkhYY.Text = mod1.HTP.Fields("GkhYY").Value '��������
htgX.txtGZH.Text = mod1.HTP.Fields("GZH").Value '�˺�
htgX.txtGyzBM.Text = mod1.HTP.Fields("GyzBM").Value '��������
htgX.txtGwiTo.Text = mod1.HTP.Fields("GwiTo").Value 'ί�д�����
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
htgX.dtpYXQ.Value = mod1.HTP.Fields("YXQ").Value '��Ч��
htgX.lblKhdh.Caption = form2Htp.txtHtbh.Text
'If Len(htgX.lblKhdh.Caption) < 5 Then
'    MsgBox ("�ú�ͬ����������,������������ϵ!")
'End If
'���²�Ʒ��
Set htgX.dtgSale.DataSource = form2Htp.adoSale

End Sub
















Public Sub wbQing() 'ά����ͬ�ֶ����
On Error Resume Next

wbHTP.txtKhmc.Text = "" '�ͻ�����
wbHTP.txtXmmc.Text = "" '��Ŀ����
wbHTP.txtKhdm.Text = "" '�ͻ�����


'��ͬǩ������
wbHTP.txtHtdate.Text = ""
wbHTP.dt1.Value = mod1.HMDa
'ά������
wbHTP.dt3.Value = mod1.HMDa
wbHTP.dt4.Value = mod1.HMDa
wbHTP.txtGLG.Text = "" '����˾
wbHTP.txtMOn.Text = "" 'ά���ʱ���
wbHTP.txtADR.Text = "" '��Ŀ��ַ
wbHTP.txtHtze.Text = "" '��ͬ�ܽ��
wbMx.txtFkBz.Text = "" '��������
wbHTP.txtCbze1.Text = "" '�ɱ��ܶ�
wbHTP.txtCbze2.Text = ""
wbHTP.txtClcb1.Text = "" '���ϳɱ�
wbHTP.txtClcb2.Text = ""
wbHTP.txtRgf1.Text = "" '�� �� ��
wbHTP.txtRGF2.Text = ""
wbHTP.txtCLF1.Text = "" '�� �� ��
wbHTP.txtCLF2.Text = ""
wbHTP.txtFbje1.Text = "" '�ְ����
wbHTP.txtFbje2.Text = ""
wbHTP.txtYf1.Text = "" '��    ��
wbHTP.txtYf2.Text = ""
wbHTP.txtYj1.Text = "" 'Ӷ    ��
wbHTP.txtYj2.Text = ""
wbHTP.txtQt1.Text = "" '��Ŀ����
wbHTP.txtQt2.Text = ""
wbHTP.txtLr1.Text = "" 'ë    ��
wbHTP.txtLr2.Text = ""
wbHTP.txtJlr1.Text = ""
wbHTP.txtJlr2.Text = ""
wbHTP.optLa.Value = False '��ֵ��Ʊ
wbHTP.optLb.Value = False '��ҵ��Ʊ
wbHTP.optLc.Value = False '����Ʊ
wbHTP.txtTc2.Text = "" '���
wbHTP.txtJy.Text = "" '������
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
wbHTP.txtTcBe.Text = 8 '��ɱ���
wbHTP.txtTcBe.Visible = False
wbHTP.lblTcBe.Visible = False
wbHTP.UpDa.Visible = False
wbHTP.txtTc2.Visible = True
wbHTP.txtTcRQ.Text = "���ȡ������"


'���wbMx��
wbMx.txtXdj.Text = "" 'Ѳ�ӵ���
wbMx.txtXgT.Text = "" 'Ѳ�ӹ�ʱ
wbMx.txtXxG.Text = "" 'Ѳ��С��
wbMx.txtJdj.Text = "" '���޵���
wbMx.txtJgT.Text = "" '���޹�ʱ
wbMx.txtJxG.Text = "" '����С��
wbMx.txtGdj.Text = "" '���̵���
wbMx.txtGgT.Text = "" '���̹�ʱ
wbMx.txtGxG.Text = "" '����С��
wbMx.txtDdj.Text = "" '���޵���
wbMx.txtDgT.Text = "" '���޹�ʱ
wbMx.txtDxG.Text = "" '����С��
wbMx.txtXgT1.Text = ""
wbMx.txtXxG1.Text = ""
wbMx.txtJgT1.Text = ""
wbMx.txtJxG1.Text = ""
wbMx.txtGgT1.Text = ""
wbMx.txtGxG1.Text = ""
wbMx.txtDgT1.Text = ""
wbMx.txtDxG1.Text = ""

wbMx.LBLhG.Caption = "" 'Ԥ���˹��ͼ�
wbMx.lblHG1.Caption = "" 'ʵ���˹��ͼ�
wbMx.cmdGzd.Caption = ""
Set wbMx.dtgGzb.DataSource = Nothing


wbMx.txtJPJE.Text = "" '������Ʊ���
wbMx.txtJPCou.Text = "" '������Ʊ����
wbMx.txtJPXG.Text = "" '������ƱС��
wbMx.txtJPXG1.Text = ""
wbMx.txtHCJE.Text = "" '������Ʊ���
wbMx.txtHCCou.Text = "" '������Ʊ����
wbMx.txtHCXG.Text = "" '������ƱС��
wbMx.txtHCXG1.Text = ""
wbMx.txtQCJE.Text = "" '��������Ʊ���
wbMx.txtQCCou.Text = "" '��������Ʊ����
wbMx.txtQCXG.Text = "" '��������ƱС��
wbMx.txtQCXG1.Text = ""
wbMx.txtZJE.Text = "" 'ס�޽��
wbMx.txtZCou.Text = "" 'ס������
wbMx.txtZXG.Text = "" 'ס��С��
wbMx.txtZXG1.Text = ""
wbMx.txtCJE.Text = "" '�ͷѽ��
wbMx.txtCCou.Text = "" '�ͷ�����
wbMx.txtCXG.Text = "" '�ͷ�С��
wbMx.txtCXG1.Text = ""
wbMx.txtDDJE.Text = "" '���س���
wbMx.txtDDCou.Text = "" '���س�������
wbMx.txtDDXG.Text = "" '���س���С��
wbMx.txtDDXG1.Text = ""


wbMx.lblCf.Caption = "" 'Ԥ�Ʋ��÷Ѻͼ�
wbMx.lblCF1.Caption = "" 'ʵ�ʲ��÷Ѻͼ�
Set wbMx.dtgCl.DataSource = Nothing

End Sub





Public Sub wbAdd() 'ά����ͬ�����������


End Sub




































































Public Sub wbBound() 'ά����ͬ�����ֶΰ�
On Error Resume Next
Dim tt As String
Dim xZ As String

'��¼����־
Call mod1.zhuDa(1, mod1.HTP.Fields("htbh").Value)
wbHTP.txtKhmc.Text = mod1.HTP.Fields("khmc").Value '�ͻ�����
wbHTP.txtXmmc.Text = mod1.HTP.Fields("xmmc").Value '��Ŀ����
wbHTP.txtYwy.Text = mod1.HTP.Fields("Ywy").Value 'ҵ��Ա
wbHTP.txtHtbh.Text = mod1.HTP.Fields("htBh").Value '��ͬ���
wbHTP.txtGLG.Text = mod1.HTP.Fields("GLG").Value '����˾
wbHTP.txtADR.Text = mod1.HTP.Fields("khADR").Value '��Ŀ��ַ
wbHTP.txtHtdate.Text = Format(mod1.HTP.Fields("htRq").Value, "Long Date") '��ͬ����
wbHTP.dt1.Value = mod1.HTP.Fields("htRq").Value
xZ = mod1.HTP.Fields("htXz").Value '��ͬ����
wbHTP.lblHid.Caption = mod1.HTP.Fields("hid").Value
Select Case xZ
'Case "A. �������ͬ"
'wbHTP.optA(0).Value = True
'Case "B1.���̺�ͬ"
'wbHTP.optA(1).Value = True
Case "C. ά����ͬ"
wbHTP.optA(3).Value = True
Case "D. ά�޺�ͬ"
wbHTP.optA(4).Value = True
'Case "E. ��Ʒ��ͬ"
'wbHTP.optA(5).Value = True
End Select
wbHTP.txtKhdm.Text = mod1.HTP.Fields("khDh").Value '�ͻ�����
wbHTP.comQy.Text = mod1.HTP.Fields("qy").Value '����

'��ͬ������ʼ��
wbHTP.dt3.Value = mod1.HTP.Fields("htQy").Value
'��ͬ���޽�����
wbHTP.dt4.Value = mod1.HTP.Fields("htQy1").Value

wbHTP.txtMOn.Text = mod1.HTP.Fields("bxQ").Value '������

If mod1.HTP.Fields("htZe").Value = 0 Then
wbHTP.txtHtze.Text = ""
Else
wbHTP.txtHtze.Text = mod1.HTP.Fields("htZe").Value '��ͬ�ܶ�
wbMx.lblHtze.Caption = wbHTP.txtHtze.Text
End If

wbMx.txtFkBz.Text = mod1.HTP.Fields("fkBz").Value '����������ע

If mod1.HTP.Fields("cbZe").Value = 0 Then
wbHTP.txtCbze1.Text = ""
Else
wbHTP.txtCbze1.Text = mod1.HTP.Fields("cbZe").Value '�ɱ��ܶ�
End If

If mod1.HTP.Fields("cbze1").Value = 0 Then
wbHTP.txtCbze2.Text = ""
Else
wbHTP.txtCbze2.Text = mod1.HTP.Fields("cbze1").Value 'ʵ�ʳɱ��ܶ�
End If


wbHTP.txtClcb1.Text = mod1.HTP.Fields("clCb").Value '���ϳɱ�



wbHTP.txtClcb2.Text = mod1.HTP.Fields("clCb1").Value 'ʵ�ʲ��ϳɱ�




wbHTP.txtRgf1.Text = mod1.HTP.Fields("rgF").Value '�˹���


wbHTP.txtRGF2.Text = mod1.HTP.Fields("rgF1").Value 'ʵ���˹���


'If mod1.HtP.Fields("clF1").Value = 0 Then '���÷�
'wbHTP.txtCLF1.Text = ""
'Else
wbHTP.txtCLF1.Text = mod1.HTP.Fields("clF1").Value
'End If
If mod1.HTP.Fields("clF2").Value = 0 Then 'ʵ�ʲ��÷�
wbHTP.txtCLF2.Text = ""
Else
wbHTP.txtCLF2.Text = mod1.HTP.Fields("clF2").Value
End If


If mod1.HTP.Fields("fbJe").Value = 0 Then
wbHTP.txtFbje1.Text = ""
Else
wbHTP.txtFbje1.Text = mod1.HTP.Fields("fbJe").Value '�ְ����
End If

If mod1.HTP.Fields("fbJe1").Value = 0 Then
wbHTP.txtFbje2.Text = ""
Else
wbHTP.txtFbje2.Text = mod1.HTP.Fields("fbJe1").Value 'ʵ�ʷְ����
End If


If mod1.HTP.Fields("yunF").Value = 0 Then
wbHTP.txtYf1.Text = ""
Else
wbHTP.txtYf1.Text = mod1.HTP.Fields("yunF").Value '�˷�
End If

If mod1.HTP.Fields("yunF1").Value = 0 Then
wbHTP.txtYf2.Text = ""
Else
wbHTP.txtYf2.Text = mod1.HTP.Fields("yunF1").Value 'ʵ���˷�
End If

'If mod1.HtP.Fields("Yj").Value = 0 Then
'wbHTP.txtYj1.Text = ""
'Else
wbHTP.txtYj1.Text = mod1.HTP.Fields("Yj").Value 'Ӷ��
'End If

If mod1.HTP.Fields("Yj1").Value = 0 Then
wbHTP.txtYj2.Text = ""
Else
wbHTP.txtYj2.Text = mod1.HTP.Fields("Yj1").Value 'ʵ��Ӷ��
End If


'If mod1.HtP.Fields("qtF").Value = 0 Then
'wbHTP.txtQt1.Text = ""
'Else
wbHTP.txtQt1.Text = mod1.HTP.Fields("qtF").Value '��Ŀ����
'End If

If mod1.HTP.Fields("qtF1").Value = 0 Then
wbHTP.txtQt2.Text = ""
Else
wbHTP.txtQt2.Text = mod1.HTP.Fields("qtF1").Value 'ʵ����Ŀ����
End If


If mod1.HTP.Fields("xmLr").Value = 0 Then
wbHTP.txtLr1.Text = ""
Else
wbHTP.txtLr1.Text = mod1.HTP.Fields("xmLr").Value '��Ŀ����
End If

If mod1.HTP.Fields("xmLr1").Value = 0 Then
wbHTP.txtLr2.Text = ""
Else
wbHTP.txtLr2.Text = mod1.HTP.Fields("xmLr1").Value 'ʵ����Ŀ����
End If

wbHTP.txtJlr1.Text = mod1.HTP.Fields("jlr1").Value
wbHTP.txtJlr2.Text = mod1.HTP.Fields("jlr2").Value

'��Ʊ����
If mod1.HTP.Fields("fpLx").Value = "��ֵ��Ʊ" Then
wbHTP.optLa.Value = True
ElseIf mod1.HTP.Fields("fpLx").Value = "��ҵ��Ʊ" Then
wbHTP.optLb.Value = True
ElseIf mod1.HTP.Fields("fpLx").Value = "����Ʊ" Then
wbHTP.optLc.Value = True
End If


If mod1.HTP.Fields("Tc1").Value = 0 Then
wbHTP.txtTc2.Text = ""
Else
wbHTP.txtTc2.Text = mod1.HTP.Fields("Tc1").Value 'ʵ�����
End If

wbHTP.txtTcBe.Text = mod1.HTP.Fields("TCbe").Value '��ɱ���


wbHTP.txtJy.Text = mod1.HTP.Fields("jy").Value '������

If mod1.HTP.Fields("ywQz").Value <> "" Then 'ҵ��Աǩ��
wbHTP.chkA.Caption = mod1.HTP.Fields("ywQz").Value
wbHTP.chkA.Tag = mod1.HTP.Fields("xywy").Value
'wbHTP.chkA.Value = 1
Else
wbHTP.chkA.Caption = ""
wbHTP.chkA.Value = 0
End If
If mod1.HTP.Fields("JzQz").Value <> "" Then '����֧��ǩ��
wbHTP.chkE.Caption = mod1.HTP.Fields("JzQz").Value
'wbHTP.chkE.Value = 1
Else
wbHTP.chkE.Caption = ""
wbHTP.chkE.Value = 0
End If
If mod1.HTP.Fields("jlQz").Value <> "" Then '���۾���ǩ��
wbHTP.chkB.Caption = mod1.HTP.Fields("jlQz").Value
wbHTP.chkB.Tag = mod1.HTP.Fields("xjlq").Value
'wbHTP.chkB.Value = 1
Else
wbHTP.chkB.Caption = ""
wbHTP.chkB.Value = 0
End If
If mod1.HTP.Fields("yzQz").Value <> "" Then '������ǩ��
wbHTP.chkC.Caption = mod1.HTP.Fields("yzQz").Value
'wbHTP.chkC.Value = 1
Else
wbHTP.chkC.Caption = ""
wbHTP.chkC.Value = 0
End If
If mod1.HTP.Fields("zjQz").Value <> "" Then '�ܾ���ǩ��
wbHTP.chkD.Caption = mod1.HTP.Fields("zjQz").Value
'wbHTP.chkD.Value = 1
Else
wbHTP.chkD.Caption = ""
wbHTP.chkD.Value = 0
End If

wbHTP.lblYw.Caption = mod1.HTP.Fields("ywDa").Value '����ǩ������
wbHTP.lblJZ.Caption = mod1.HTP.Fields("JzDa").Value '����֧��ǩ������
wbHTP.lblJl.Caption = mod1.HTP.Fields("JlDa").Value '���۾���ǩ������
wbHTP.lblYZ.Caption = mod1.HTP.Fields("YzDa").Value '������ǩ������
wbHTP.lblZJ.Caption = mod1.HTP.Fields("ZjDa").Value '�ܾ���ǩ������
wbHTP.txtXMNr.Text = mod1.HTP.Fields("xmnr").Value '

If mod1.HTP.Fields("jTf").Value = True Then
    wbHTP.cmdCount.Caption = "�ѽ���"
    wbHTP.cmdCount.Enabled = False
Else
    wbHTP.cmdCount.Caption = "����"
    wbHTP.cmdCount.Enabled = True
  
End If


If IsNull(mod1.HTP.Fields("TCRQ").Value) = False Then
    wbHTP.txtTcRQ.Text = mod1.HTP.Fields("TCRQ").Value '���ȡ������
End If

'���Ϊ�ɺ�ͬ,������׶���Ϊ��ɫ
If mod1.HTP.Fields("XGG").Value = True Then
    wbHTP.optP.BackColor = &HC0FFFF
ElseIf mod1.HTP.Fields("XGG").Value = 0 Then
    wbHTP.optP.BackColor = &H8000000F
End If

wbHTP.optP.Enabled = False
wbHTP.optG.Enabled = False
wbHTP.optZ.Enabled = False
wbHTP.optW.Enabled = False

'��ִͬ�з�
If mod1.HTP.Fields(51).Value = 0 Then
    wbHTP.optP.Value = True
'    wbHTP.optP.Enabled = True
'    '�����ǩ����,����Ϊ�����ߴ�,����Խ��и���
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

''�����ͬ����û��ȫͨ�����߲���С�⣬�򡰺�ִͬ�з��ֶβ��ܱ༭
'If wbHTP.chkD.Caption <> "" And frmLogin.Combo1.Text = "��ӱ" Then
'wbHTP.frmZt.Enabled = True
'Else

'End If

'��wbRGMX��
Set adoRGF = CreateObject("adodb.recordset")
tt = "Select * from wbRGMX where htBh='" & wbHTP.txtHtbh.Text & "'"
adoRGF.Close
adoRGF.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
wbMx.txtXdj.Text = adoRGF.Fields("Xdj").Value 'Ѳ�ӵ���
wbMx.txtXgT.Text = adoRGF.Fields("XgT").Value 'Ѳ�ӹ�ʱ
wbMx.txtXxG.Text = adoRGF.Fields("XxG").Value 'Ѳ��С��
wbMx.txtJdj.Text = adoRGF.Fields("Jdj").Value '���޵���
wbMx.txtJgT.Text = adoRGF.Fields("JgT").Value '���޹�ʱ
wbMx.txtJxG.Text = adoRGF.Fields("JxG").Value '����С��
wbMx.txtGdj.Text = adoRGF.Fields("Gdj").Value '���̵���
wbMx.txtGgT.Text = adoRGF.Fields("GgT").Value '���̹�ʱ
wbMx.txtGxG.Text = adoRGF.Fields("GxG").Value '����С��
wbMx.txtDdj.Text = adoRGF.Fields("Ddj").Value '���޵���
wbMx.txtDgT.Text = adoRGF.Fields("DgT").Value '���޹�ʱ
wbMx.txtDxG.Text = adoRGF.Fields("DxG").Value '����С��
'Ԥ���˹��ͼ�
wbMx.LBLhG.Caption = wbHTP.txtRgf1.Text
wbMx.txtJPJE.Text = adoRGF.Fields("JPJE").Value '������Ʊ���
wbMx.txtJPCou.Text = adoRGF.Fields("JPCou").Value '������Ʊ����
wbMx.txtJPXG.Text = adoRGF.Fields("JPXG").Value '������ƱС��
wbMx.txtHCJE.Text = adoRGF.Fields("HCJE").Value '������Ʊ���
wbMx.txtHCCou.Text = adoRGF.Fields("HCCou").Value '������Ʊ����
wbMx.txtHCXG.Text = adoRGF.Fields("HCXG").Value '������ƱС��
wbMx.txtQCJE.Text = adoRGF.Fields("QCJE").Value '��������Ʊ���
wbMx.txtQCCou.Text = adoRGF.Fields("QCCou").Value '��������Ʊ����
wbMx.txtQCXG.Text = adoRGF.Fields("QCXG").Value '��������ƱС��
wbMx.txtZJE.Text = adoRGF.Fields("ZJE").Value 'ס�޽��
wbMx.txtZCou.Text = adoRGF.Fields("ZCou").Value 'ס������
wbMx.txtZXG.Text = adoRGF.Fields("ZXG").Value 'ס��С��
wbMx.txtCJE.Text = adoRGF.Fields("CJE").Value '�ͷѽ��
wbMx.txtCCou.Text = adoRGF.Fields("CCou").Value '�ͷ�����
wbMx.txtCXG.Text = adoRGF.Fields("CXG").Value '�ͷ�С��
wbMx.txtDDJE.Text = adoRGF.Fields("DDJE").Value '���س���
wbMx.txtDDCou.Text = adoRGF.Fields("DDCou").Value '���س�������
wbMx.txtDDXG.Text = adoRGF.Fields("DDXG").Value '���س���С��
'Ԥ�Ʋ��÷Ѻͼ�
wbMx.lblCf.Caption = wbHTP.txtCLF1.Text

'��Ӷ���
tt = "Select * from Yongjin where htBh='" & wbHTP.txtHtbh.Text & "' order by yId"
frmYJ.adoYj.Recordset.Close
frmYJ.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Set frmYJ.dtgYJ.DataSource = frmYJ.adoYj

wbHTP.frmZt.Enabled = True
End Sub






Public Sub GXZJ(Htbh As String)  '���㹺����ͬ��ʵ�ʳɱ��ܶʵ����������
Dim tt As String
On Error Resume Next
'����ʵ�ʳɱ��ܶ���ϳɱ�+Ԥ������+�˷�+Ӷ��
tt = "update htping set cbZe1=clcb1+qtF1+yunF1+Yj1 where htbh='" & Htbh & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'����ʵ������
tt = "update htping set xmLr1=htZe-cbZe1 where htbh='" & Htbh & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'����ʵ�����
tt = "update htping set Tc1=xmLr1*0.08 where htbh='" & Htbh & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
End Sub

Public Sub gxMr() 'Ĭ�ϵĹ�����ͬ����
htgX.txtT2.Text = "������������׼"
'htgX.txtT4.Text = "��������"
htgX.txtT5.Text = "�跽����"
htgX.txtT7.Text = "��׼��װ"
htgX.txtT8.Text = "�����ֳ����������������������������"
'htgX.txtT10.Text = "�����"
htgX.txtT12.Text = "���ա��л����񹲺͹���ͬ����"
htgX.txtT13.Text = "Э�̡��ٲá����ϣ��Ϻ���"
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
'frmWbNew.txtZe.Text = mod1.HTP.Fields("htbh").Value  '�����տ�
'frmWbNew.txtEd.Text = mod1.HTP.Fields("htbh").Value
'��Ʊ����

If mod1.HTP.Fields("fpLX").Value = "��ֵ��Ʊ" Then
    frmWbNew.optLa.Value = True
ElseIf mod1.HTP.Fields("fpLX").Value = "��ҵ��Ʊ" Then
    frmWbNew.optLb.Value = True
ElseIf mod1.HTP.Fields("fpLX").Value = "����Ʊ" Then
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
frmWbNew.txtQt2 = mod1.HTP.Fields("qtF").Value '�Ѿ���������Ŀ����


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

'��������
'ҵ��
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

tt = "select rq as ����,je as ���,bz as ��ע,mid from htpingJt where hid=" & Val(frmWbNew.lblHid.Caption) & " and delf=1 order by mid desc"
mod1.mJt.Close
mod1.mJt.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.mJt.RecordCount) = True Then
    MsgBox ("��ȡ���ݴ���2.2!")
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
    MsgBox ("��ȡ���ݴ���2.3!")
    Exit Sub
End If
frmWbNew.txtJTf.Text = mod1.HTP.Fields("je").Value

tt = "select rq as ����,je as ���,bz as ��ע,mid from htpingQk where hid=" & Val(frmWbNew.lblHid.Caption) & " and delf=1 order by mid desc"
mod1.mYjF.Close
mod1.mYjF.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.mYjF.RecordCount) = True Then
    MsgBox ("��ȡ���ݴ���2.8!")
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
    MsgBox ("��ȡ���ݴ���2.6!")
    Exit Sub
End If
frmWbNew.txtYjf.Text = mod1.HTP.Fields("je").Value



tt = "SELECT rp_dd as ����,amtn_cls as ���,rem as ��ע FROM TF_MON where rp_id=1 and cas_no='" & frmWbNew.txtHtbh.Text & "' order by rp_dd"
mod1.mQk.Close
mod1.mQk.Open tt, mod1.workTx, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.mQk.RecordCount) = True Then
    MsgBox ("��ȡ���ݴ���2.5!")
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
    MsgBox ("��ȡ���ݴ���2.6!")
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



If frmWbNew.lblHtxz.Caption = "ά��" Then
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
        tt = "select jzpb as ����Ʒ��,jzxh as �����ͺ�,sl as ����,jxId from wbjb where baoid=" & Val(frmWbNew.lblBaoId.Caption)
        Set frmWbNew.adoA = CreateObject("adodb.recordset")
        frmWbNew.adoA.Close
        frmWbNew.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmWbNew.dtgA.DataSource = frmWbNew.adoA
        frmWbNew.dtgA.Visible = True
        frmWbNew.cmdTk.Visible = True
    Else
        frmWbNew.cmdTk.Visible = False
        frmWbNew.dtgA.Visible = False
        '�걣
        tt = "select * from xunJIaWbView where wbx='�걣' and bid=" & mod1.HTP.Fields("bid").Value
        frmWbNew.adoWb.Close
        frmWbNew.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmWbNew.dtgWb.DataSource = frmWbNew.adoWb
        frmWbNew.dtgWb.FixedRows = 0
        frmWbNew.dtgWb.MergeCol(1) = True
        frmWbNew.dtgWb.MergeCol(2) = True
        frmWbNew.dtgWb.MergeCol(3) = True
        frmWbNew.dtgWb.MergeCells = 3
        frmWbNew.dtgWb.FixedRows = 1
        '�����
        tt = "select * from xunJIaWbView where wbx='����' and bid=" & mod1.HTP.Fields("bid").Value
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
    

    
    '��ʾ��Ʒ�б�
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
    '��ʾ�ɱ���
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
ElseIf frmWbNew.lblHtxz.Caption = "����" Or frmWbNew.lblHtxz.Caption = "���̷ְ�" Then
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
        tt = "select jzpb as ����Ʒ��,jzxh as �����ͺ�,sl as ����,jxId from wbjb where baoid=" & Val(frmWbNew.lblBaoId.Caption)
        Set frmWbNew.adoA = CreateObject("adodb.recordset")
        frmWbNew.adoA.Close
        frmWbNew.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmWbNew.dtgA.DataSource = frmWbNew.adoA
        frmWbNew.dtgA.Visible = True
    Else
        frmWbNew.dtgA.Visible = False
    End If
    
    '��ʾ��Ʒ�б�
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
    '��ʾ�ɱ���
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
ElseIf frmWbNew.lblHtxz.Caption = "��Ʒ" Or frmWbNew.lblHtxz.Caption = "�����" Or frmWbNew.lblHtxz.Caption = "���̷ְ�" Or frmWbNew.lblHtxz.Caption = "����" Then
    '��ʾ��Ʒ�б�
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
    '��ʾ�ɱ���
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

'��ʾ�̶�����
tt = "select lb as �������,year(nd) as ���,qdj as ����,rl as ����,xg as С��,baoid,hid,gid from xmgd where hid=" & Val(frmWbNew.lblHid.Caption)
frmWbNew.adoGD.Close
frmWbNew.adoGD.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWbNew.dtgGD.DataSource = frmWbNew.adoGD
tt = "select sum(xg) as xg from xmgd where baoid=" & Val(frmWbNew.lblBaoId.Caption)
frmWbNew.adoHGD.Close
frmWbNew.adoHGD.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
frmWbNew.txtGd.Text = frmWbNew.adoHGD.Fields("xg").Value
'frmWbNew.txtXm2.Text = Val(frmWbNew.txtGd.Text) + Val(frmWbNew.txtXm.Text)




'��Ӧ�տ��
tt = "select * from htFk where htbh='" & frmWbNew.lblHid.Caption & "'"
frmWbNew.adoFk.Close
frmWbNew.adoFk.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWbNew.dtgFk.DataSource = frmWbNew.adoFk
'��Ӷ���
tt = "select yED as �տ���,YingFu as ֧�����,yid,ywy from yongjin where htbh='" & frmWbNew.txtHtbh.Text & "' order by yid"
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
If mod1.DName = "������" Or mod1.DName = "����" Then
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
    mod1.HTP.MoveNext '��һ�����鰴ť�������,����,������һ��¼
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
                If frmWbNew.cmdQm(oo).Caption = "�Ͼ��쾭��" Then
                    frmWbNew.cmdQm(oo).Caption = "�Ͼ��쾭��"
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
