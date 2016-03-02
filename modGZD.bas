Attribute VB_Name = "modGZD"

Public Sub gzd1Qing()
Dim oo As Integer
For oo = 1 To 17
    NewGZD1.BA(oo).Text = ""
Next
For oo = 1 To 60
    NewGZD1.TA(oo).Text = ""
Next
For oo = 1 To 63
    NewGZD1.C1(oo).Value = 0
Next
NewGZD1.FPA.Value = False
NewGZD1.FPB.Value = False
NewGZD1.FPC.Value = False
NewGZD1.FPD.Value = False
NewGZD1.cmdQm(0).Caption = ""
NewGZD1.lblTm(0).Caption = ""
NewGZD1.lblGid.Caption = ""
NewGZD1.lblBh.Caption = ""
NewGZD1.cmdQm(0).Caption = ""
NewGZD1.lblKhdh.Caption = ""
NewGZD1.comhtBh.Visible = False
NewGZD1.comXmmc.Visible = False
NewGZD1.dtgRen.Visible = False
NewGZD1.tabNr.Tab = 0
End Sub
Public Sub gzd2Qing()
Dim oo As Integer
For oo = 1 To 17
    NewGzd2.BA(oo).Text = ""
Next
For oo = 1 To 73
    NewGzd2.TA(oo).Text = ""
Next
For oo = 1 To 80
    NewGzd2.C1(oo).Value = 0
Next
NewGzd2.FPA.Value = False
NewGzd2.FPB.Value = False
NewGzd2.FPC.Value = False
NewGzd2.FPD.Value = False
NewGzd2.cmdQm(0).Caption = ""
NewGzd2.lblTm(0).Caption = ""
NewGzd2.lblGid.Caption = ""
NewGzd2.lblBh.Caption = ""
NewGzd2.cmdQm(0).Caption = ""
NewGzd2.lblKhdh.Caption = ""
NewGzd2.comhtBh.Visible = False
NewGzd2.comXmmc.Visible = False
NewGzd2.dtgRen.Visible = False
NewGzd2.tabNr.Tab = 0
End Sub

Public Sub gzd3Qing()
Dim oo As Integer
For oo = 1 To 17
    NewGzd3.BA(oo).Text = ""
Next
For oo = 1 To 27
    NewGzd3.TA(oo).Text = ""
Next
For oo = 1 To 162
    NewGzd3.C1(oo).Value = 0
Next
NewGzd3.FPA.Value = False
NewGzd3.FPB.Value = False
NewGzd3.FPC.Value = False
NewGzd3.FPD.Value = False
NewGzd3.cmdQm(0).Caption = ""
NewGzd3.lblTm(0).Caption = ""
NewGzd3.lblGid.Caption = ""
NewGzd3.lblBh.Caption = ""
NewGzd3.cmdQm(0).Caption = ""
NewGzd3.lblKhdh.Caption = ""
NewGzd3.comhtBh.Visible = False
NewGzd3.comXmmc.Visible = False
NewGzd3.dtgRen.Visible = False
NewGzd3.tabNr.Tab = 0
End Sub

Public Sub gzd4Qing()
Dim oo As Integer
For oo = 1 To 17
    NewGzd4.BA(oo).Text = ""
Next
For oo = 1 To 27
    NewGzd4.TA(oo).Text = ""
Next
For oo = 1 To 156
    NewGzd4.C1(oo).Value = 0
Next
NewGzd4.FPA.Value = False
NewGzd4.FPB.Value = False
NewGzd4.FPC.Value = False
NewGzd4.FPD.Value = False
NewGzd4.cmdQm(0).Caption = ""
NewGzd4.lblTm(0).Caption = ""
NewGzd4.lblGid.Caption = ""
NewGzd4.lblBh.Caption = ""
NewGzd4.cmdQm(0).Caption = ""
NewGzd4.lblKhdh.Caption = ""
NewGzd4.comhtBh.Visible = False
NewGzd4.comXmmc.Visible = False
NewGzd4.dtgRen.Visible = False
NewGzd4.tabNr.Tab = 0
End Sub
Public Sub gzd5Qing()
Dim oo As Integer
For oo = 1 To 17
    NewGzd5.BA(oo).Text = ""
Next
For oo = 1 To 110
    NewGzd5.TA(oo).Text = ""
Next
For oo = 1 To 22
    NewGzd5.C1(oo).Value = 0
Next
NewGzd5.FPA.Value = False
NewGzd5.FPB.Value = False
NewGzd5.FPC.Value = False
NewGzd5.FPD.Value = False
NewGzd5.cmdQm(0).Caption = ""
NewGzd5.lblTm(0).Caption = ""
NewGzd5.lblGid.Caption = ""
NewGzd5.lblBh.Caption = ""
NewGzd5.cmdQm(0).Caption = ""
NewGzd5.lblKhdh.Caption = ""
NewGzd5.comhtBh.Visible = False
NewGzd5.comXmmc.Visible = False
NewGzd5.dtgRen.Visible = False
NewGzd5.tabNr.Tab = 0
End Sub
Public Sub gzd6Qing()
Dim oo As Integer
For oo = 1 To 17
    NewGzd6.BA(oo).Text = ""
Next
For oo = 1 To 108
    NewGzd6.TA(oo).Text = ""
Next
For oo = 1 To 32
    NewGzd6.C1(oo).Value = 0
Next
NewGzd6.FPA.Value = False
NewGzd6.FPB.Value = False
NewGzd6.FPC.Value = False
NewGzd6.FPD.Value = False
NewGzd6.cmdQm(0).Caption = ""
NewGzd6.lblTm(0).Caption = ""
NewGzd6.lblGid.Caption = ""
NewGzd6.lblBh.Caption = ""
NewGzd6.cmdQm(0).Caption = ""
NewGzd6.lblKhdh.Caption = ""
NewGzd6.comhtBh.Visible = False
NewGzd6.comXmmc.Visible = False
NewGzd6.dtgRen.Visible = False
NewGzd6.tabNr.Tab = 0
End Sub

Public Sub gzd1Bound(Gid As Long)
Dim tt As String
Dim oo As Integer
On Error Resume Next
tt = "select * from newgzd where gid=" & Gid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
For oo = 1 To 17
    NewGZD1.BA(oo).Text = mod1.HTP.Fields("a" & oo).Value
Next
NewGZD1.BA(6).Text = Format(NewGZD1.BA(6).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGZD1.BA(14).Text = Format(NewGZD1.BA(14).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGZD1.BA(16).Text = Format(NewGZD1.BA(16).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
For oo = 1 To 60
    NewGZD1.TA(oo).Text = mod1.HTP.Fields("mat" & oo).Value
Next
For oo = 1 To 63
    'NewGZD1.C1(oo).Value = mod1.HTP.Fields("mac" & oo).Value
    If mod1.HTP.Fields("mac" & oo).Value = True Then
        NewGZD1.C1(oo).Value = 1
    End If
Next
If mod1.HTP.Fields("fp").Value = 1 Then
    NewGZD1.FPA.Value = True
ElseIf mod1.HTP.Fields("fp").Value = 2 Then
    NewGZD1.FPB.Value = True
ElseIf mod1.HTP.Fields("fp").Value = 3 Then
    NewGZD1.FPC.Value = True
ElseIf mod1.HTP.Fields("fp").Value = 4 Then
    NewGZD1.FPD.Value = True
End If
NewGZD1.cmdQm(0).Caption = mod1.HTP.Fields("ywy").Value
NewGZD1.lblTm(0).Caption = mod1.HTP.Fields("trq").Value
NewGZD1.LBLKjj.Visible = False
If IsNull(mod1.HTP.Fields("trq").Value) = True Then
    NewGZD1.cmdQm(0).Caption = ""
    NewGZD1.LBLKjj.Visible = True
End If
NewGZD1.lblGid.Caption = Gid
NewGZD1.lblBh.Caption = mod1.HTP.Fields("bh").Value
NewGZD1.lblKhdh.Caption = mod1.HTP.Fields("khdh").Value
End Sub
Public Sub gzd2Bound(Gid As Long)
Dim tt As String
Dim oo As Integer
On Error Resume Next
tt = "select * from newgzd where gid=" & Gid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
For oo = 1 To 17
    NewGzd2.BA(oo).Text = mod1.HTP.Fields("a" & oo).Value
Next
NewGzd2.BA(6).Text = Format(NewGzd2.BA(6).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGzd2.BA(14).Text = Format(NewGzd2.BA(14).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGzd2.BA(16).Text = Format(NewGzd2.BA(16).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
For oo = 1 To 73
    NewGzd2.TA(oo).Text = mod1.HTP.Fields("mat" & oo).Value
Next
For oo = 1 To 80
    If mod1.HTP.Fields("mac" & oo).Value = True Then
        NewGzd2.C1(oo).Value = 1
    End If
Next
If mod1.HTP.Fields("fp").Value = 1 Then
    NewGzd2.FPA.Value = True
ElseIf mod1.HTP.Fields("fp").Value = 2 Then
    NewGzd2.FPB.Value = True
ElseIf mod1.HTP.Fields("fp").Value = 3 Then
    NewGzd2.FPC.Value = True
ElseIf mod1.HTP.Fields("fp").Value = 4 Then
    NewGzd2.FPD.Value = True
End If
NewGzd2.cmdQm(0).Caption = mod1.HTP.Fields("ywy").Value
NewGzd2.lblTm(0).Caption = mod1.HTP.Fields("trq").Value
NewGzd2.LBLKjj.Visible = False
If IsNull(mod1.HTP.Fields("trq").Value) = True Then
    NewGzd2.cmdQm(0).Caption = ""
    NewGzd2.LBLKjj.Visible = True
End If
NewGzd2.lblGid.Caption = Gid
NewGzd2.lblBh.Caption = mod1.HTP.Fields("bh").Value
NewGzd2.lblKhdh.Caption = mod1.HTP.Fields("khdh").Value
End Sub

Public Sub gzd3Bound(Gid As Long)
Dim tt As String
Dim oo As Integer
On Error Resume Next
tt = "select * from newgzd where gid=" & Gid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
For oo = 1 To 17
    NewGzd3.BA(oo).Text = mod1.HTP.Fields("a" & oo).Value
Next
NewGzd3.BA(6).Text = Format(NewGzd3.BA(6).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGzd3.BA(14).Text = Format(NewGzd3.BA(14).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGzd3.BA(16).Text = Format(NewGzd3.BA(16).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
For oo = 1 To 27
    NewGzd3.TA(oo).Text = mod1.HTP.Fields("mat" & oo).Value
Next
For oo = 1 To 162
    If mod1.HTP.Fields("mac" & oo).Value = True Then
        NewGzd3.C1(oo).Value = 1
    End If
Next
If mod1.HTP.Fields("fp").Value = 1 Then
    NewGzd3.FPA.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 2 Then
    NewGzd3.FPB.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 3 Then
    NewGzd3.FPC.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 4 Then
    NewGzd3.FPD.Value = False
End If
NewGzd3.cmdQm(0).Caption = mod1.HTP.Fields("ywy").Value
NewGzd3.lblTm(0).Caption = mod1.HTP.Fields("trq").Value
NewGzd3.LBLKjj.Visible = False
If IsNull(mod1.HTP.Fields("trq").Value) = True Then
    NewGzd3.cmdQm(0).Caption = ""
    NewGzd3.LBLKjj.Visible = True
End If
NewGzd3.lblGid.Caption = Gid
NewGzd3.lblBh.Caption = mod1.HTP.Fields("bh").Value
NewGzd3.lblKhdh.Caption = mod1.HTP.Fields("khdh").Value
End Sub

Public Sub gzd4Bound(Gid As Long)
Dim tt As String
Dim oo As Integer
On Error Resume Next
tt = "select * from newgzd where gid=" & Gid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
For oo = 1 To 17
    NewGzd4.BA(oo).Text = mod1.HTP.Fields("a" & oo).Value
Next
NewGzd4.BA(6).Text = Format(NewGzd4.BA(6).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGzd4.BA(14).Text = Format(NewGzd4.BA(14).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGzd4.BA(16).Text = Format(NewGzd4.BA(16).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
For oo = 1 To 27
    NewGzd4.TA(oo).Text = mod1.HTP.Fields("mat" & oo).Value
Next
For oo = 1 To 156
    If mod1.HTP.Fields("mac" & oo).Value = True Then
        NewGzd4.C1(oo).Value = 1
    End If
Next
If mod1.HTP.Fields("fp").Value = 1 Then
    NewGzd4.FPA.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 2 Then
    NewGzd4.FPB.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 3 Then
    NewGzd4.FPC.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 4 Then
    NewGzd4.FPD.Value = False
End If
NewGzd4.cmdQm(0).Caption = mod1.HTP.Fields("ywy").Value
NewGzd4.lblTm(0).Caption = mod1.HTP.Fields("trq").Value
NewGzd4.LBLKjj.Visible = False
If IsNull(mod1.HTP.Fields("trq").Value) = True Then
    NewGzd4.cmdQm(0).Caption = ""
    NewGzd4.LBLKjj.Visible = True
End If
NewGzd4.lblGid.Caption = Gid
NewGzd4.lblBh.Caption = mod1.HTP.Fields("bh").Value
NewGzd4.lblKhdh.Caption = mod1.HTP.Fields("khdh").Value
End Sub
Public Sub gzd5Bound(Gid As Long)
Dim tt As String
Dim oo As Integer
On Error Resume Next
tt = "select * from newgzd where gid=" & Gid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
For oo = 1 To 17
    NewGzd5.BA(oo).Text = mod1.HTP.Fields("a" & oo).Value
Next
NewGzd5.BA(6).Text = Format(NewGzd5.BA(6).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGzd5.BA(14).Text = Format(NewGzd5.BA(14).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGzd5.BA(16).Text = Format(NewGzd5.BA(16).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
For oo = 1 To 110
    NewGzd5.TA(oo).Text = mod1.HTP.Fields("mat" & oo).Value
Next
For oo = 1 To 22
    If mod1.HTP.Fields("mac" & oo).Value = True Then
        NewGzd5.C1(oo).Value = 1
    End If
Next
If mod1.HTP.Fields("fp").Value = 1 Then
    NewGzd5.FPA.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 2 Then
    NewGzd5.FPB.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 3 Then
    NewGzd5.FPC.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 4 Then
    NewGzd5.FPD.Value = False
End If
NewGzd5.cmdQm(0).Caption = mod1.HTP.Fields("ywy").Value
NewGzd5.lblTm(0).Caption = mod1.HTP.Fields("trq").Value
NewGzd5.LBLKjj.Visible = False
If IsNull(mod1.HTP.Fields("trq").Value) = True Then
    NewGzd5.cmdQm(0).Caption = ""
    NewGzd5.LBLKjj.Visible = True
End If
NewGzd5.lblGid.Caption = Gid
NewGzd5.lblBh.Caption = mod1.HTP.Fields("bh").Value
NewGzd5.lblKhdh.Caption = mod1.HTP.Fields("khdh").Value
End Sub
Public Sub gzd6Bound(Gid As Long)
Dim tt As String
Dim oo As Integer
On Error Resume Next
tt = "select * from newgzd where gid=" & Gid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
For oo = 1 To 17
    NewGzd6.BA(oo).Text = mod1.HTP.Fields("a" & oo).Value
Next
NewGzd6.BA(6).Text = Format(NewGzd6.BA(6).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGzd6.BA(14).Text = Format(NewGzd6.BA(14).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGzd6.BA(16).Text = Format(NewGzd6.BA(16).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
For oo = 1 To 108
    NewGzd6.TA(oo).Text = mod1.HTP.Fields("mat" & oo).Value
Next
For oo = 1 To 32
    If mod1.HTP.Fields("mac" & oo).Value = True Then
        NewGzd6.C1(oo).Value = 1
    End If
Next
If mod1.HTP.Fields("fp").Value = 1 Then
    NewGzd6.FPA.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 2 Then
    NewGzd6.FPB.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 3 Then
    NewGzd6.FPC.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 4 Then
    NewGzd6.FPD.Value = False
End If
NewGzd6.cmdQm(0).Caption = mod1.HTP.Fields("ywy").Value
NewGzd6.lblTm(0).Caption = mod1.HTP.Fields("trq").Value
NewGzd6.LBLKjj.Visible = False
If IsNull(mod1.HTP.Fields("trq").Value) = True Then
    NewGzd6.cmdQm(0).Caption = ""
    NewGzd6.LBLKjj.Visible = True
End If
NewGzd6.lblGid.Caption = Gid
NewGzd6.lblBh.Caption = mod1.HTP.Fields("bh").Value
NewGzd6.lblKhdh.Caption = mod1.HTP.Fields("khdh").Value
End Sub
Public Sub gzd7Bound(Gid As Long)
Dim tt As String
Dim oo As Integer
On Error Resume Next
tt = "select * from newgzd where gid=" & Gid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
For oo = 1 To 17
    NewGzd7.BA(oo).Text = mod1.HTP.Fields("a" & oo).Value
Next
NewGzd7.BA(6).Text = Format(NewGzd7.BA(6).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGzd7.BA(14).Text = Format(NewGzd7.BA(14).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGzd7.BA(16).Text = Format(NewGzd7.BA(16).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
For oo = 1 To 64
    NewGzd7.TA(oo).Text = mod1.HTP.Fields("mat" & oo).Value
Next
For oo = 1 To 4
    If mod1.HTP.Fields("mac" & oo).Value = True Then
        NewGzd7.C1(oo).Value = 1
    End If
Next
If mod1.HTP.Fields("fp").Value = 1 Then
    NewGzd7.FPA.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 2 Then
    NewGzd7.FPB.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 3 Then
    NewGzd7.FPC.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 4 Then
    NewGzd7.FPD.Value = False
End If
NewGzd7.cmdQm(0).Caption = mod1.HTP.Fields("ywy").Value
NewGzd7.lblTm(0).Caption = mod1.HTP.Fields("trq").Value
NewGzd7.LBLKjj.Visible = False
If IsNull(mod1.HTP.Fields("trq").Value) = True Then
    NewGzd7.cmdQm(0).Caption = ""
    NewGzd7.LBLKjj.Visible = True
End If
NewGzd7.lblGid.Caption = Gid
NewGzd7.lblBh.Caption = mod1.HTP.Fields("bh").Value
NewGzd7.lblKhdh.Caption = mod1.HTP.Fields("khdh").Value
End Sub

Public Sub gzd8Bound(Gid As Long)
Dim tt As String
Dim oo As Integer
On Error Resume Next
tt = "select * from newgzd where gid=" & Gid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
For oo = 1 To 17
    NewGZD8.BA(oo).Text = mod1.HTP.Fields("a" & oo).Value
Next
NewGZD8.BA(6).Text = Format(NewGZD8.BA(6).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGZD8.BA(14).Text = Format(NewGZD8.BA(14).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
NewGZD8.BA(16).Text = Format(NewGZD8.BA(16).Text, "YYYY/MM/DD", vbUseSystemDayOfWeek)
For oo = 1 To 62
    NewGZD8.TA(oo).Text = mod1.HTP.Fields("mat" & oo).Value
Next
For oo = 1 To 38
    If mod1.HTP.Fields("mac" & oo).Value = True Then
        NewGZD8.C1(oo).Value = 1
    End If
Next
If mod1.HTP.Fields("fp").Value = 1 Then
    NewGZD8.FPA.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 2 Then
    NewGZD8.FPB.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 3 Then
    NewGZD8.FPC.Value = False
ElseIf mod1.HTP.Fields("fp").Value = 4 Then
    NewGZD8.FPD.Value = False
End If
NewGZD8.cmdQm(0).Caption = mod1.HTP.Fields("ywy").Value
NewGZD8.lblTm(0).Caption = mod1.HTP.Fields("trq").Value
NewGZD8.LBLKjj.Visible = False
If IsNull(mod1.HTP.Fields("trq").Value) = True Then
    NewGZD8.cmdQm(0).Caption = ""
    NewGZD8.LBLKjj.Visible = True
End If
NewGZD8.lblGid.Caption = Gid
NewGZD8.lblBh.Caption = mod1.HTP.Fields("bh").Value
NewGZD8.lblKhdh.Caption = mod1.HTP.Fields("khdh").Value
End Sub
Public Sub gzd7Qing()
Dim oo As Integer
For oo = 1 To 17
    NewGzd7.BA(oo).Text = ""
Next
For oo = 1 To 64
    NewGzd7.TA(oo).Text = ""
Next
For oo = 1 To 4
    NewGzd7.C1(oo).Value = 0
Next
NewGzd7.FPA.Value = False
NewGzd7.FPB.Value = False
NewGzd7.FPC.Value = False
NewGzd7.FPD.Value = False
NewGzd7.cmdQm(0).Caption = ""
NewGzd7.lblTm(0).Caption = ""
NewGzd7.lblGid.Caption = ""
NewGzd7.lblBh.Caption = ""
NewGzd7.cmdQm(0).Caption = ""
NewGzd7.lblKhdh.Caption = ""
NewGzd7.comhtBh.Visible = False
NewGzd7.comXmmc.Visible = False
NewGzd7.dtgRen.Visible = False
NewGzd7.tabNr.Tab = 0
End Sub

Public Sub gzd8Qing()
Dim oo As Integer
For oo = 1 To 17
    NewGZD8.BA(oo).Text = ""
Next
For oo = 1 To 62
    NewGZD8.TA(oo).Text = ""
Next
For oo = 1 To 38
    NewGZD8.C1(oo).Value = 0
Next
NewGZD8.FPA.Value = False
NewGZD8.FPB.Value = False
NewGZD8.FPC.Value = False
NewGZD8.FPD.Value = False
NewGZD8.cmdQm(0).Caption = ""
NewGZD8.lblTm(0).Caption = ""
NewGZD8.lblGid.Caption = ""
NewGZD8.lblBh.Caption = ""
NewGZD8.cmdQm(0).Caption = ""
NewGZD8.lblKhdh.Caption = ""
NewGZD8.comhtBh.Visible = False
NewGZD8.comXmmc.Visible = False
'NewGZD8.dtgRen.Visible = False
NewGZD8.tabNr.Tab = 0
End Sub

