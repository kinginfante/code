Attribute VB_Name = "modGGL"
Option Explicit
'Public Gid As Integer
Public Oid As Double '当前打开的公告栏编号

Public Sub zTing() '霓虹灯字停
'frmGGL.Picture = LoadPicture(App.Path & "\pic\公告栏e.jpg")
frmGGL.Picture = frmGGL.pic1.Picture
frmGGL.Timer1.Enabled = False
frmGGL.Timer2.Enabled = False
frmGGL.Timer3.Enabled = False
frmGGL.Timer4.Enabled = False
frmGGL.Timer5.Enabled = False
frmGGL.lblA.Visible = False
frmGGL.lblB.Visible = False
frmGGL.lblC.Visible = False
frmGGL.lblD.Visible = False
frmGGL.lblE.Visible = False
If frmGGL.comLb.Text = "晨会类" Then
    frmGGL.Timer1.Enabled = True
    frmGGL.Timer2.Enabled = True
    frmGGL.Timer3.Enabled = True
    frmGGL.Timer4.Enabled = True
    frmGGL.Timer5.Enabled = True
End If
End Sub

Public Sub GGLBound()  '公告栏数据获取
Dim tt As String
On Error Resume Next
'Call modGGL.CHZT




'tt = "Select gnr,zz,rq,gid,fdx,wzid from ggl where " & mod1.DName & "=0 order by gid desc"
tt = "Select top 1 gid from ggl where (" & mod1.DName & "=0 or lb='胡萝卜' and " & mod1.DName & " is null) order by gid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
'Oid = mod1.HTP.Fields("gid").Value

If mod1.HTP.RecordCount > 0 Then
    Set frmGGL.adoGGl = CreateObject("adodb.recordset")
    'tt = "Select top 1 gnr,zz,rq,gid,fdx,wzid,lb,fid from ggl where " & mod1.DName & "=0 and gid<" & Oid & " order by gid desc"
    tt = "Select top 1 gnr,zz,rq,gid,fdx,wzid,lb,fid from ggl where (" & mod1.DName & "=0 or " & mod1.DName & " is null and lb='胡萝卜') order by gid desc"
    frmGGL.adoGGl.Close
    frmGGL.adoGGl.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
    Oid = frmGGL.adoGGl.Fields("gid").Value
    frmGGL.Gid = Oid
    frmGGL.rihNr.Text = frmGGL.adoGGl.Fields("Gnr").Value
    If Left(frmGGL.adoGGl.Fields("zz").Value, 1) = "n" Then
        frmGGL.lblZZ.Caption = "匿名者"
    Else
        frmGGL.lblZZ.Caption = frmGGL.adoGGl.Fields("zz").Value
    End If

    frmGGL.lblDate.Caption = frmGGL.adoGGl.Fields("rq").Value
    If IsNull(frmGGL.adoGGl.Fields("lb").Value) = True Then
        frmGGL.comLb.Visible = False
        frmGGL.lblLb.Visible = False
    Else
        frmGGL.comLb.Text = frmGGL.adoGGl.Fields("lb").Value
        frmGGL.comLb.Visible = True
        frmGGL.lblLb.Visible = True
        frmGGL.comLb.Locked = True
        frmGGL.frmLx.Enabled = False

    End If
    frmGGL.Show
    frmGGL.WindowState = 0
    frmGGL.ZOrder 0
    'frmZu.Enabled = False

    '判断字颜色
    frmGGL.rihNr.SelStart = 0
    frmGGL.rihNr.SelLength = Len(frmGGL.rihNr.Text)

        frmGGL.rihNr.SelColor = &HFF0000

    frmGGL.rihNr.SelFontSize = frmGGL.adoGGl.Fields("Fdx").Value
    frmGGL.rihNr.SelStart = 0
    frmGGL.rihNr.SelLength = 0
End If

If frmGGL.lblZZ.Caption = mod1.DName Or mod1.DName = "马晓聪" Then
frmGGL.cmdDel.Enabled = True
Else
frmGGL.cmdDel.Enabled = False
End If


frmGGL.cmdYjb.Visible = False
'If IsNull(frmGGL.adoGG.Recordset.Fields("wzid").Value) = False Then
'
'
'    If Left(frmGGL.rihNr.Text, 3) = "请注意" Then
'        frmGGL.cmdYjb.Visible = True
'    Else
'        frmGGL.cmdXQ.Visible = True
'    End If
'End If
frmGGL.WindowState = 0
frmGGL.ZOrder 0
tt = "Select top 1 username,userid from worker where qy='KKK'"
frmGGL.adoRen.Recordset.Close
frmGGL.adoRen.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

    frmGGL.timNew.Enabled = True
End Sub


Public Sub GGLR()
Dim tt As String
Dim Fid As Long
On Error Resume Next


frmGGL.cmdSave.Enabled = False

    Set frmGGL.adoGGl = CreateObject("adodb.recordset")
    If frmGGL.cmdXJ.Caption = "已看" Then
        If mod1.Mname = "马晓聪" Then
            tt = "SELECT dbo.NGGL.gnr, dbo.NGGL.ZUid, dbo.NGGL.rq, dbo.NGGL.LB, dbo.NGGL.QF, " & _
                "dbo.NGGL.Gid, dbo.NGGLDetail.Uid, dbo.NGGLDetail.LKF, dbo.NGGLFile.FName,dbo.NGGLFile.Fid " & _
                    "FROM dbo.NGGL LEFT OUTER JOIN dbo.NGGLFile ON dbo.NGGL.Gid = dbo.NGGLFile.gid LEFT OUTER JOIN " & _
                    "dbo.NGGLDetail ON dbo.NGGL.Gid = dbo.NGGLDetail.Gid"
        Else
            tt = "Select top 1 gnr,zz,rq,gid,fdx,wzid,lb,fid, " & mod1.DName & " from ggl where  gid<" & Oid & " and  (" & mod1.DName & "=0 or " & mod1.DName & " is null and lb='胡萝卜') order by gid desc"
        End If
    Else
        If mod1.Mname = "马晓聪" Then
            tt = "Select top 1 gnr,zz,rq,gid,fdx,wzid,lb,fid, " & mod1.DName & " from ggl where  gid<" & Oid & " and " & mod1.DName & " =1 order by gid desc"
        Else
            tt = "Select top 1 gnr,zz,rq,gid,fdx,wzid,lb,fid, " & mod1.DName & " from ggl where  gid<" & Oid & " and " & mod1.DName & " =1 order by gid desc"
        End If
    End If
    frmGGL.adoGGl.Close
    frmGGL.adoGGl.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
    If frmGGL.adoGGl.RecordCount = 1 Then
        Oid = frmGGL.adoGGl.Fields("gid").Value
        frmGGL.rihNr.Text = frmGGL.adoGGl.Fields("Gnr").Value
        If Left(frmGGL.adoGGl.Fields("zz").Value, 1) = "n" Then
            frmGGL.lblZZ.Caption = "匿名者"
        Else
            frmGGL.lblZZ.Caption = frmGGL.adoGGl.Fields("zz").Value
        End If
        frmGGL.lblDate.Caption = frmGGL.adoGGl.Fields("rq").Value
        Fid = frmGGL.adoGGl.Fields("fid").Value
        If IsNull(frmGGL.adoGGl.Fields("lb").Value) = True Then
            frmGGL.comLb.Visible = False
            frmGGL.lblLb.Visible = False
        Else
            frmGGL.comLb.Text = frmGGL.adoGGl.Fields("lb").Value
            frmGGL.comLb.Visible = True
            frmGGL.lblLb.Visible = True
            frmGGL.comLb.Locked = True
            frmGGL.frmLx.Enabled = False

        End If
        frmGGL.Show
        frmGGL.WindowState = 0
        frmGGL.ZOrder 0
        
        frmZu.Enabled = False
        
        '判断字颜色
        frmGGL.rihNr.SelStart = 0
        frmGGL.rihNr.SelLength = Len(frmGGL.rihNr.Text)
    
        If frmGGL.adoGGl.Fields(mod1.DName).Value = False Then
            frmGGL.rihNr.SelColor = &HFF0000
        Else
            frmGGL.rihNr.SelColor = &H80000012
        End If
    
        frmGGL.rihNr.SelFontSize = frmGGL.adoGGl.Fields("Fdx").Value
        frmGGL.rihNr.SelStart = 0
        frmGGL.rihNr.SelLength = 0
    'End If
    
        If frmGGL.lblZZ.Caption = mod1.DName Or mod1.DName = "马晓聪" Then
        frmGGL.cmdDel.Enabled = True
        Else
        frmGGL.cmdDel.Enabled = False
        End If
        

        frmGGL.cmdYjb.Visible = False
        
        If IsNull(frmGGL.adoGG.Recordset.Fields("wzid").Value) = False Then


            If Left(frmGGL.rihNr.Text, 3) = "请注意" Then
                frmGGL.cmdYjb.Visible = True
            Else

            End If
        End If
        frmGGL.cmdZx.Enabled = True
        frmGGL.cmdReply.Enabled = True
        frmGGL.frmRen.Visible = False
        frmGGL.cmdPre.Enabled = True
    Else
        frmGGL.cmdNext.Enabled = False
    End If
    frmGGL.cmdYjb.Visible = False

If frmGGL.adoGGl.RecordCount = 1 Then
    If IsNull(frmGGL.adoGGl.Fields("wzid").Value) = False Then
    
    
        If Left(frmGGL.rihNr.Text, 3) = "请注意" Then
            frmGGL.cmdYjb.Visible = True
        Else

        End If
    End If
End If



End Sub

Public Sub GGLL()
 Dim tt As String
 Dim Fid As Long
On Error Resume Next


frmGGL.cmdSave.Enabled = False

    Set frmGGL.adoGGl = CreateObject("adodb.recordset")
    If frmGGL.cmdXJ.Caption = "已看" Then
        'tt = "Select top 1 gnr,zz,rq,gid,fdx,wzid,lb,fid, " & mod1.DName & " from ggl where  gid>" & Oid & " and not(" & mod1.DName & " =1) order by gid"
        'tt = "Select top 1 gnr,zz,rq,gid,fdx,wzid,lb,fid, " & mod1.DName & " from ggl where  gid>" & Oid & " and " & mod1.DName & " =0 order by gid"
        tt = "Select top 1 gnr,zz,rq,gid,fdx,wzid,lb,fid, " & mod1.DName & " from ggl where  gid>" & Oid & " and  (" & mod1.DName & "=0 or " & mod1.DName & " is null and lb='胡萝卜') order by gid"
    Else
        tt = "Select top 1 gnr,zz,rq,gid,fdx,wzid,lb,fid, " & mod1.DName & " from ggl where  gid>" & Oid & " and " & mod1.DName & " =1 order by gid"
        'tt = "Select top 1 gnr,zz,rq,gid,fdx,wzid,lb,fid, " & mod1.DName & " from ggl where (" & mod1.DName & "=0 or " & mod1.DName & " is null and lb='胡萝卜') order by gid desc"
    End If
    frmGGL.adoGGl.Close
    frmGGL.adoGGl.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
    If frmGGL.adoGGl.RecordCount = 1 Then
        Oid = frmGGL.adoGGl.Fields("gid").Value
        frmGGL.rihNr.Text = frmGGL.adoGGl.Fields("Gnr").Value
        If Left(frmGGL.adoGGl.Fields("zz").Value, 1) = "n" Then
            frmGGL.lblZZ.Caption = "匿名者"
        Else
            frmGGL.lblZZ.Caption = frmGGL.adoGGl.Fields("zz").Value
        End If
        frmGGL.lblDate.Caption = frmGGL.adoGGl.Fields("rq").Value
        Fid = frmGGL.adoGGl.Fields("fid").Value
        If IsNull(frmGGL.adoGGl.Fields("lb").Value) = True Then
            frmGGL.comLb.Visible = False
            frmGGL.lblLb.Visible = False
        Else
            frmGGL.comLb.Text = frmGGL.adoGGl.Fields("lb").Value
            frmGGL.comLb.Visible = True
            frmGGL.lblLb.Visible = True
            frmGGL.comLb.Locked = True
            frmGGL.frmLx.Enabled = False

        End If
        frmGGL.Show
        frmGGL.WindowState = 0
        frmGGL.ZOrder 0
        frmZu.Enabled = False
        
        '判断字颜色
        frmGGL.rihNr.SelStart = 0
        frmGGL.rihNr.SelLength = Len(frmGGL.rihNr.Text)
    
        If frmGGL.adoGGl.Fields(mod1.DName).Value = False Then
            frmGGL.rihNr.SelColor = &HFF0000
        Else
            frmGGL.rihNr.SelColor = &H80000012
        End If
    
        frmGGL.rihNr.SelFontSize = frmGGL.adoGGl.Fields("Fdx").Value
        frmGGL.rihNr.SelStart = 0
        frmGGL.rihNr.SelLength = 0
    'End If
    
        If frmGGL.lblZZ.Caption = mod1.DName Or mod1.DName = "马晓聪" Then
        frmGGL.cmdDel.Enabled = True
        Else
        frmGGL.cmdDel.Enabled = False
        End If
        

        frmGGL.cmdYjb.Visible = False
        
        If IsNull(frmGGL.adoGGl.Fields("wzid").Value) = False Then


            If Left(frmGGL.rihNr.Text, 3) = "请注意" Then
                frmGGL.cmdYjb.Visible = True
            Else

            End If
        End If
        frmGGL.cmdZx.Enabled = True
        frmGGL.cmdReply.Enabled = True
        frmGGL.frmRen.Visible = False
        frmGGL.cmdNext.Enabled = True
    Else
        frmGGL.cmdPre.Enabled = False
    End If
    frmGGL.cmdYjb.Visible = False

If frmGGL.adoGGl.RecordCount = 1 Then
    If IsNull(frmGGL.adoGGl.Fields("wzid").Value) = False Then
    
    
        If Left(frmGGL.rihNr.Text, 3) = "请注意" Then
            frmGGL.cmdYjb.Visible = True
        Else

        End If
    End If
End If




End Sub

Public Sub CHZ() '晨会彩字.
frmGGL.Timer1.Enabled = True
frmGGL.Timer2.Enabled = True
frmGGL.Timer3.Enabled = True
frmGGL.Timer4.Enabled = True
frmGGL.Timer5.Enabled = True
frmGGL.frmCa.Visible = True
frmGGL.frmCb.Visible = True

End Sub
Public Sub CHZT() '晨会彩字.
frmGGL.Timer1.Enabled = False
frmGGL.Timer2.Enabled = False
frmGGL.Timer3.Enabled = False
frmGGL.Timer4.Enabled = False
frmGGL.Timer5.Enabled = False
frmGGL.frmCa.Visible = False
frmGGL.frmCb.Visible = False

End Sub
