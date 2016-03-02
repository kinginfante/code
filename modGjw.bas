Attribute VB_Name = "modGjw"

Public GDL(0 To 20) As Long
Public Sub GjwQing()
On Error Resume Next
Dim oo As Integer
For oo = 1 To 15
    frmGJW.liD(oo).X1 = 4695
    frmGJW.liD(oo).x2 = 4770
    frmGJW.liD(oo).BorderColor = &H80000008
    frmGJW.liD(oo).Visible = False
    frmGJW.txtNr(oo).Text = ""
Next
frmGJW.txtDate.Text = ""
For oo = 0 To 31
    frmGJW.txtDay(oo).Text = ""
Next
frmGJW.txtXmmc.Text = ""
frmGJW.txtHtbh.Text = ""
frmGJW.txtZu.Text = ""
frmGJW.txtBid.Text = ""
For oo = 0 To 4
    frmGJW.lblQM(oo).Caption = ""
    frmGJW.lblTm(oo).Caption = ""
    frmGJW.cmdQm(oo).Caption = ""
Next

'坐标
GDL(0) = 4440
GDL(1) = 4926
GDL(2) = 5385
GDL(3) = 5835
GDL(4) = 6270
GDL(5) = 6750
GDL(6) = 7230
GDL(7) = 7665
GDL(8) = 8115
GDL(9) = 8550
GDL(10) = 9030
GDL(11) = 9480
GDL(12) = 9945
GDL(13) = 10395
GDL(14) = 10860
GDL(15) = 11310
GDL(16) = 11790
GDL(17) = 12225
GDL(18) = 12690
GDL(19) = 13140
GDL(20) = 13650
frmGJW.cmdMod.Enabled = True
frmGJW.cmdSave.Enabled = False

End Sub

Public Sub GjwOpen(Gid As Long)
Dim tt As String
Dim oo As Integer
On Error Resume Next
tt = "select * from gjw where gid=" & Gid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
frmGJW.txtXmmc.Text = mod1.HTP.Fields("xmmc").Value
frmGJW.txtHtbh.Text = mod1.HTP.Fields("htbh").Value
frmGJW.txtZu.Text = mod1.HTP.Fields("zname").Value
frmGJW.txtDate.Text = mod1.HTP.Fields("Gdate").Value
frmGJW.txtBid.Text = mod1.HTP.Fields("bid").Value
For oo = 0 To 20
    frmGJW.txtDay(oo).ToolTipText = mod1.HTP.Fields("gday" & oo).Value
    If IsNull(mod1.HTP.Fields("gday" & oo).Value) = False Then
        frmGJW.txtDay(oo).Text = Day(mod1.HTP.Fields("gday" & oo).Value) & "日"
    End If
Next
frmGJW.LblTrq.Caption = mod1.HTP.Fields("trq").Value
frmGJW.lblFwid.Caption = mod1.HTP.Fields("fwid").Value
frmGJW.lblGid.Caption = mod1.HTP.Fields("gid").Value
frmGJW.lblLc.Caption = mod1.HTP.Fields("lc").Value
frmGJW.lblLcRen.Caption = mod1.HTP.Fields("lcren").Value
frmGJW.lblLcUid.Caption = mod1.HTP.Fields("lcuid").Value
frmGJW.lblYwy.Caption = mod1.HTP.Fields("zname").Value
frmGJW.lblUid.Caption = mod1.HTP.Fields("uid").Value
frmGJW.LblTrq.Caption = mod1.HTP.Fields("trq").Value
frmGJW.lblPwf.Caption = mod1.HTP.Fields("pwf").Value
frmGJW.lblGid.Caption = mod1.HTP.Fields("GID").Value

'显示进度线
For oo = 1 To 15
    frmGJW.txtNr(oo).Text = mod1.HTP.Fields("nr" & oo).Value
    If frmGJW.txtNr(oo).Text <> "" Then
        frmGJW.liD(oo).Visible = True
        frmGJW.liD(oo).X1 = mod1.HTP.Fields("x1" & oo).Value
        frmGJW.liD(oo).x2 = mod1.HTP.Fields("x2" & oo).Value
        frmGJW.liD(oo).BorderColor = mod1.HTP.Fields("lcolor" & oo).Value
        
    End If
Next


'显示签字按钮
tt = "select * from qmrz where qdbh='" & frmGJW.lblGid.Caption & "' and btz=51 order by zid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
mod1.HTP.MoveFirst
oo = 0
Do While Not mod1.HTP.EOF
    frmGJW.lblQM(oo).Caption = mod1.HTP.Fields("qlabel").Value
    frmGJW.cmdQm(oo).Caption = mod1.HTP.Fields("qren").Value
    frmGJW.lblTm(oo).Caption = mod1.HTP.Fields("qrq").Value
    mod1.HTP.MoveNext
    oo = oo + 1
Loop
End Sub

