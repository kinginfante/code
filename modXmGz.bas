Attribute VB_Name = "modXmGz"
Option Explicit
Public Gid As Integer
Public FR As Date 'һ�ܵĵ�һ��
Public LR As Date 'һ�ܵ����һ��
Public Ti As Boolean '�Ƿ�Ϊ����ӵļ�¼
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
    mod1.HTP.MoveNext '��һ�����鰴ť�������,����,������һ��¼
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
dayWeek = "��"
Case 2
dayWeek = "һ"
Case 3
dayWeek = "��"
Case 4
dayWeek = "��"
Case 5
dayWeek = "��"
Case 6
dayWeek = "��"
Case 7
dayWeek = "��"
End Select
End Function


Public Sub xmAdd() '��Ŀ�������
Dim Pd As Integer
Dim tt As String
On Error Resume Next
tt = "Select * from xmGz where gid=" & frmGzNr.lblGid.Caption
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
If mod1.HTP.RecordCount < 1 Then
    MsgBox "����,�벻Ҫ�رճ���,����������ϵ��"
    Exit Sub
End If
mod1.HTP.Update "xid", Val(frmGzNr.lblXid.Caption)
mod1.HTP.Update "khQc", frmGzNr.lblXmmc.Caption  '�ͻ�ȫ��


mod1.HTP.Update "xmAdr", frmGzNr.lblAdr.Caption '��Ŀ��ַ
mod1.HTP.Update "BfMd", Left(frmGzNr.txtBfMd.Text, 500) '�ݷ�Ŀ��
mod1.HTP.Update "BfJg", Left(frmGzNr.txtBfJg.Text, 500) '�ݷý��
mod1.HTP.Update "xbCC", Left(frmGzNr.txtXBCC.Text, 500) '�²��ƻ�
'mod1.htp.Update "XDBZ", Left(frmGzNr.txtXDBZ.Text, 50) '�ж�����
mod1.HTP.Update "aTime", frmGzNr.lblRq.Caption 'ʱ�䰲��
mod1.HTP.Update "XmFy", Val(frmGzNr.txtXmFy.Text)  '��Ŀ����
mod1.HTP.Update "xM", Left(frmGzNr.txtXm.Text, 500) '��Ŀ����
mod1.HTP.Update "jzDC", Left(frmGzNr.txtjzDC.Text, 500) '��������
mod1.HTP.Update "zgPd", Left(frmGzNr.txtzgPd.Text, 500) '��������
mod1.HTP.Update "zgQz", frmGzNr.lblZGQZ.Caption '����ǩ��
mod1.HTP.Update "Lb", 1 '���

mod1.HTP.UpdateBatch

'������ñ�
frmGzNr.adoFy.Recordset.UpdateBatch

'�ͻ�ƽ̨
If frmGzNr.optA.Value = True Then
    Pd = 0
ElseIf frmGzNr.optB.Value = True Then
    Pd = 30
ElseIf frmGzNr.optC.Value = True Then
    Pd = 60
ElseIf frmGzNr.optD.Value = True Then
    Pd = 90
End If

'������Ŀ���ϱ��е���Ŀƽ̨
tt = "update xmzl set khJb=" & Pd & " where xid=" & Val(frmGzNr.lblXid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
End Sub

Public Sub jiAdd() '�ƻ����
Dim tt As String
On Error Resume Next
tt = "Select * from xmGz where gid=" & modXmGz.Gid
frmGzJ.adoXmgz.Recordset.Close
frmGzJ.adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
frmGzJ.adoXmgz.Recordset.Update "khDh", frmGzJ.lblDm.Caption '�ͻ�����
frmGzJ.adoXmgz.Recordset.Update "khQc", frmGzJ.lblKhmc.Caption   '�ͻ�ȫ��
frmGzJ.adoXmgz.Recordset.Update "xmmc", frmGzJ.lblKhmc.Caption
'�ͻ�ƽ̨
If frmGzJ.optA.Value = True Then
frmGzJ.adoXmgz.Recordset.Update "khJb", 0
ElseIf frmGzJ.optB.Value = True Then
frmGzJ.adoXmgz.Recordset.Update "khJb", 30
ElseIf frmGzJ.optC.Value = True Then
frmGzJ.adoXmgz.Recordset.Update "khJb", 60
ElseIf frmGzJ.optD.Value = True Then
frmGzJ.adoXmgz.Recordset.Update "khJb", 90
End If

frmGzJ.adoXmgz.Recordset.Update "xmAdr", frmGzJ.lblAdr.Caption '��Ŀ��ַ
frmGzJ.adoXmgz.Recordset.Update "BfMd", Left(frmGzJ.txtBfMd.Text, 500) '�ݷ�Ŀ��
'frmGzNr.adoXmgz.Recordset.Update "BfJg", Left(frmGzNr.txtBfJg.Text, 50) '�ݷý��
frmGzJ.adoXmgz.Recordset.Update "XDBZ", Left(frmGzJ.txtXDBZ.Text, 500) '�ж�����
frmGzJ.adoXmgz.Recordset.Update "aTime", frmGzJ.lblRq.Caption 'ʱ�䰲��
'frmgzJ.adoXmgz.Recordset.Update "XmFy", Val(frmgzJ.txtXmFy.Text)  '��Ŀ����
frmGzJ.adoXmgz.Recordset.Update "xM", Left(frmGzJ.txtXm.Text, 500) '��Ŀ����
frmGzJ.adoXmgz.Recordset.Update "jzDC", Left(frmGzJ.txtjzDC.Text, 500) '��������
frmGzJ.adoXmgz.Recordset.Update "zgPd", Left(frmGzJ.txtzgPd.Text, 500) '��������
frmGzJ.adoXmgz.Recordset.Update "zgQz", frmGzJ.lblZGQZ.Caption '����ǩ��
frmGzJ.adoXmgz.Recordset.Update "Lb", 0 '���
frmGzJ.adoXmgz.Recordset.UpdateBatch

''������ñ�
'frmGzNr.adoFy.Recordset.UpdateBatch
End Sub



Public Sub xmBound(Gid As Long)  '��Ŀ���ٰ�
Dim tt As String
On Error Resume Next
tt = "Select * from xmGz where gid=" & Gid
frmGzNr.adoXmgz.Recordset.Close
frmGzNr.adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'frmGzNr.lblDm.Caption = frmGzNr.adoXmgz.Recordset.Fields("khDh").Value '�ͻ�����
frmGzNr.lblXid.Caption = frmGzNr.adoXmgz.Recordset.Fields("xid").Value
frmGzNr.lblXmmc.Caption = frmGzNr.adoXmgz.Recordset.Fields("xmmc").Value '��Ŀȫ��
frmGzNr.comKhmc.Text = frmGzNr.adoXmgz.Recordset.Fields("khqc").Value '�ͻ�����
frmGzNr.lblYwy.Caption = frmGzNr.adoXmgz.Recordset.Fields("ywy").Value 'ҵ��Ա
frmGzNr.lblGid.Caption = frmGzNr.adoXmgz.Recordset.Fields("gid").Value




frmGzNr.lblAdr.Caption = frmGzNr.adoXmgz.Recordset.Fields("xmAdr").Value '��Ŀ��ַ
frmGzNr.txtBfMd.Text = frmGzNr.adoXmgz.Recordset.Fields("BfMd").Value '�ݷ�Ŀ��
frmGzNr.txtBfJg.Text = frmGzNr.adoXmgz.Recordset.Fields("BfJg").Value '�ݷý��
frmGzNr.txtXBCC.Text = frmGzNr.adoXmgz.Recordset.Fields("xbCC").Value '�²��ƻ�
'frmGzNr.txtXDBZ.Text = frmGzNr.adoXmgz.Recordset.Fields("XDBZ").Value '�ж�����
frmGzNr.lblRq.Caption = frmGzNr.adoXmgz.Recordset.Fields("aTime").Value 'ʱ�䰲��
frmGzNr.txtXmFy.Text = frmGzNr.adoXmgz.Recordset.Fields("XmFy").Value '��Ŀ����
frmGzNr.txtXm.Text = frmGzNr.adoXmgz.Recordset.Fields("xM").Value '��Ŀ����
frmGzNr.txtjzDC.Text = frmGzNr.adoXmgz.Recordset.Fields("jzDC").Value '��������
frmGzNr.txtzgPd.Text = frmGzNr.adoXmgz.Recordset.Fields("zgPd").Value '��������
frmGzNr.lblZGQZ.Caption = frmGzNr.adoXmgz.Recordset.Fields("zgQz").Value '����ǩ��

frmGzNr.lblLc.Caption = frmGzNr.adoXmgz.Recordset.Fields("Lc").Value
frmGzNr.lblLcRen.Caption = Trim(frmGzNr.adoXmgz.Recordset.Fields("LcRen").Value)
frmGzNr.lblLcUid.Caption = Trim(frmGzNr.adoXmgz.Recordset.Fields("LcUid").Value)
frmGzNr.lblFwid.Caption = frmGzNr.adoXmgz.Recordset.Fields("Fwid").Value
frmGzNr.lblNlb.Caption = frmGzNr.adoXmgz.Recordset.Fields("Nlb").Value
frmGzNr.lblLcou.Caption = frmGzNr.adoXmgz.Recordset.Fields("Lcou").Value
frmGzNr.lblHtbh.Caption = frmGzNr.adoXmgz.Recordset.Fields("htbh").Value

'�󶨷��ñ�
tt = "Select * from fyTg where gid=" & Val(frmGzNr.lblGid.Caption)
frmGzNr.adoFy.Recordset.Close
frmGzNr.adoFy.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
Set frmGzNr.dtgFy.DataSource = frmGzNr.adoFy

'���¿ͻ������������
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
    
'ȡ����Ŀ�ܷ���
'������Ŀ�ܷ���
tt = "select sum(xg) from fybx where khmc='" & frmGzNr.lblXmmc.Caption & "' and not(qrq is null)"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
frmGzNr.lblCxmFy.Caption = mod1.HTP.Fields(0).Value

        frmGzNr.optC.Enabled = False
        frmGzNr.optD.Enabled = False
'�ͻ�ƽ̨
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
    
    Call modXmGz.OpenXMGZAN(True) '�򿪰�ť
    frmGzNr.cmdSave.Enabled = False
End Sub
Public Sub OpenXMGZAN(LX As Boolean)
Dim tt As String
Dim oo As Integer
On Error Resume Next
If LX = True Then   '�����ռ�
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
Public Sub jiBound() '�ƻ���
Dim tt As String
On Error Resume Next
tt = "Select * from xmGz where gid=" & modXmGz.Gid
frmGzJ.adoXmgz.Recordset.Close
frmGzJ.adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
frmGzJ.lblDm.Caption = frmGzJ.adoXmgz.Recordset.Fields("khDh").Value '�ͻ�����
frmGzJ.lblKhmc.Caption = frmGzJ.adoXmgz.Recordset.Fields("xmmc").Value '�ͻ�ȫ��
frmGzJ.lblYwy.Caption = frmGzJ.adoXmgz.Recordset.Fields("ywy").Value 'ҵ��Ա

'�ͻ�ƽ̨
If frmGzJ.adoXmgz.Recordset.Fields("khJb").Value = 0 Then
    frmGzJ.optA.Value = True
ElseIf frmGzJ.adoXmgz.Recordset.Fields("khJb").Value = 30 Then
    frmGzJ.optB.Value = True
ElseIf frmGzJ.adoXmgz.Recordset.Fields("khJb").Value = 60 Then
    frmGzJ.optC.Value = True
ElseIf frmGzJ.adoXmgz.Recordset.Fields("khJb").Value = 90 Then
    frmGzJ.optD.Value = True
End If

frmGzJ.lblAdr.Caption = frmGzJ.adoXmgz.Recordset.Fields("xmAdr").Value '��Ŀ��ַ
frmGzJ.txtBfMd.Text = frmGzJ.adoXmgz.Recordset.Fields("BfMd").Value '�ݷ�Ŀ��
'frmgzJ.txtBfJg.Text = frmgzJ.adoXmgz.Recordset.Fields("BfJg").Value '�ݷý��
'frmgzJ.txtXBCC.Text = frmgzJ.adoXmgz.Recordset.Fields("xbCC").Value '�²��ƻ�
frmGzJ.txtXDBZ.Text = frmGzJ.adoXmgz.Recordset.Fields("XDBZ").Value '�ж�����
frmGzJ.lblRq.Caption = frmGzJ.adoXmgz.Recordset.Fields("aTime").Value 'ʱ�䰲��
'frmGzJ.txtXmFy.Text = frmGzJ.adoXmgz.Recordset.Fields("XmFy").Value '��Ŀ����
frmGzJ.txtXm.Text = frmGzJ.adoXmgz.Recordset.Fields("xM").Value '��Ŀ����
frmGzJ.txtjzDC.Text = frmGzJ.adoXmgz.Recordset.Fields("jzDC").Value '��������
frmGzJ.txtzgPd.Text = frmGzJ.adoXmgz.Recordset.Fields("zgPd").Value '��������
frmGzJ.lblZGQZ.Caption = frmGzJ.adoXmgz.Recordset.Fields("zgQz").Value '����ǩ��

''�󶨷��ñ�
'tt = "Select * from fyTg where gid=" & modXmGz.Gid
'frmGzJ.adoFy.Recordset.Close
'frmGzJ.adoFy.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'Set frmGzJ.dtgFy.DataSource = frmGzJ.adoFy

End Sub

Public Sub xmQing() '��Ŀ�����ֶ����
frmGzNr.lblXid.Caption = "" '�ͻ�����
frmGzNr.lblXmmc.Caption = "" '�ͻ�ȫ��
frmGzNr.comKhmc.Text = ""
frmGzNr.lblAdr.Caption = ""
frmGzNr.lblKid.Caption = ""
frmGzNr.lblGid.Caption = ""
frmGzNr.lblYwy.Caption = ""

'�ͻ�ƽ̨
    frmGzNr.optA.Value = False
    frmGzNr.optB.Value = False
    frmGzNr.optC.Value = False
    frmGzNr.optD.Value = False
frmGzNr.txtJw.Text = "" '�ͻ��������
frmGzNr.lblAdr.Caption = "" '��Ŀ��ַ
frmGzNr.txtBfMd.Text = ""  '�ݷ�Ŀ��
frmGzNr.txtBfJg.Text = ""  '�ݷý��
'frmGzNr.txtXDBZ.Text = "" '�ж�����
frmGzNr.txtXBCC.Text = "" '�²���ʩ
frmGzNr.lblRq.Caption = ""  'ʱ�䰲��
frmGzNr.txtXmFy.Text = "" '��Ŀ����
frmGzNr.lblCxmFy.Caption = "" '��Ŀ�ܷ���
frmGzNr.txtXm.Text = ""  '��Ŀ����
frmGzNr.txtjzDC.Text = ""  '��������
frmGzNr.txtzgPd.Text = ""  '��������
frmGzNr.lblZGQZ.Caption = ""  '����ǩ��

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
frmGzNr.lblHtbh.Caption = "" '���α����������ĺ�ͬ���
Set frmGzNr.dtgRen.DataSource = Nothing

End Sub

Public Sub jhQing() '�ƻ��ֶ����
frmGzJ.lblDm.Caption = "" '�ͻ�����
'frmGzJ.lblXmmc.Caption = "" '�ͻ�ȫ��

'�ͻ�ƽ̨
    frmGzJ.optA.Value = False
    frmGzJ.optB.Value = False
    frmGzJ.optC.Value = False
    frmGzJ.optD.Value = False

frmGzJ.lblAdr.Caption = "" '��Ŀ��ַ
frmGzJ.txtBfMd.Text = ""  '�ݷ�Ŀ��
'frmGzJ.txtBfJg.Text = ""  '�ݷý��
frmGzJ.txtXDBZ.Text = "" '�ж�����
frmGzJ.lblRq.Caption = ""  'ʱ�䰲��
'frmGzJ.txtXmFy.Text = "" '��Ŀ����
frmGzJ.txtXm.Text = ""  '��Ŀ����
frmGzJ.txtjzDC.Text = ""  '��������
frmGzJ.txtzgPd.Text = ""  '��������
frmGzJ.lblZGQZ.Caption = ""  '����ǩ��

End Sub



Public Sub FyQing() 'Ӫ�������������

    frmFYBX.lblBh.Caption = ""
    frmFYBX.comQy.Caption = "�Ϻ�"
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








Public Sub XMPwf() '��Ŀ����ϸ��
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
