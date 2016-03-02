VERSION 5.00
Begin VB.Form frmHyxz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "请选择行业性质"
   ClientHeight    =   315
   ClientLeft      =   5445
   ClientTop       =   4095
   ClientWidth     =   3435
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   315
   ScaleWidth      =   3435
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定"
      Height          =   315
      Left            =   2730
      TabIndex        =   1
      Top             =   0
      Width           =   705
   End
   Begin VB.ComboBox comHyxz 
      Height          =   300
      ItemData        =   "frmHyxz.frx":0000
      Left            =   0
      List            =   "frmHyxz.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "frmHyxz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

End Sub

Private Sub cmdOk_Click()
Dim bt As String
Dim tt As String
On Error Resume Next
    If comHyxz.Text = "" Then Exit Sub
    '先取得代号编码
    tt = "Select * from Xohyxz where hyxz='" & comHyxz.Text & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    bt = mod1.HTP.Fields("Bm").Value

    tt = "Select max(khDh) as cou from khzl where khDh like '%" & bt & "%'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    If IsNull(mod1.HTP.Fields("cou").Value) = True Then
    wbDN.txtKhdm.Text = bt & Format(1, "0000")
    Else
    wbDN.txtKhdm.Text = bt & Format(Val(Right(mod1.HTP.Fields("cou").Value, 4)) + 1, "0000")
    End If
    wbDN.tabKh.Enabled = True
    
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "khjia"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@Uid") = mod1.DHid
    mod1.cmd.Parameters("@khdh") = wbDN.txtKhdm.Text
    mod1.cmd.Parameters("@bm") = mod1.Bm
    mod1.cmd.Parameters("@hyxz") = comHyxz.Text
    mod1.cmd.Parameters("@Lcou") = Right(frmKhBr.cmdNew.ToolTipText, 1) '流程总数
    mod1.cmd.Parameters("@Lcou") = 3
    mod1.cmd.Parameters("@Lc") = 0 '当前流程
    mod1.cmd.Parameters("@lcRen") = mod1.DName
    mod1.cmd.Parameters("@lcUid") = mod1.DHid
    mod1.cmd.Parameters("@xid") = Val(wbDN.lblXid.Caption)
    mod1.cmd.Parameters("@nLb") = 88
    mod1.cmd.Parameters("@rid") = 0
    mod1.cmd.Parameters("@kid") = 0
    mod1.cmd.Execute

    wbDN.comXyxz.Text = comHyxz.Text
    
    wbDN.lblKid.Caption = mod1.cmd.Parameters("@kid").Value
    If wbDN.optYz.Value = True Then
        wbDN.lblYZ.Tag = mod1.cmd.Parameters("@kid").Value
    ElseIf wbDN.lblQT(1).Value = True Then
        wbDN.lblQT(1).Tag = mod1.cmd.Parameters("@kid").Value
    ElseIf wbDN.lblQT(2).Value = True Then
        wbDN.lblQT(2).Tag = mod1.cmd.Parameters("@kid").Value
    ElseIf wbDN.lblQT(3).Value = True Then
        wbDN.lblQT(3).Tag = mod1.cmd.Parameters("@kid").Value
    ElseIf wbDN.lblQT(4).Value = True Then
        wbDN.lblQT(4).Tag = mod1.cmd.Parameters("@kid").Value
    ElseIf wbDN.lblQT(5).Value = True Then
        wbDN.lblQT(5).Tag = mod1.cmd.Parameters("@kid").Value
    End If
    wbDN.lblRid.Caption = mod1.cmd.Parameters("@rid").Value
    wbDN.lblYwy.Caption = mod1.DName
    wbDN.lblUid.Caption = mod1.DHid
    wbDN.lblLcRen.Caption = mod1.DName
    wbDN.lblLcUid.Caption = mod1.DHid
    wbDN.lblXywy.Caption = mod1.DName
    wbDN.lblXuid.Caption = mod1.DHid
    
    Set mod1.cmd = Nothing
    If wbDN.optWy.Value = True Then '物业公司
        tt = "Select * from khloPan where kid='" & wbDN.lblKid.Caption & "'"
        wbDN.adoLouPan.Recordset.Close
        wbDN.adoLouPan.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        wbDN.frmGL.Visible = True
        wbDN.frmJz.Visible = False
        wbDN.cmdLadd.Enabled = False
    Else
        wbDN.frmGL.Visible = False
        wbDN.frmJz.Visible = True
    End If
    

    
    '设置流程按钮
    Call mod1.khLcBut(38)
    frmHyxz.Visible = False
    wbDN.Enabled = True
    wbDN.tabKh.Tab = 0
    wbDN.cmdMod.Enabled = False
    wbDN.cmdSave.Enabled = True
    
'新添加的客户,可以编辑

'wbDN.txtKhmc.Locked = True
wbDN.comXZ.Locked = False '企业性质
wbDN.comXyxz.Locked = False '行业性质
wbDN.txtAdr1.Locked = False '项目地址
wbDN.comQy.Locked = False '区域
wbDN.txtFH.Locked = False '国税号
wbDN.txtKhYY.Locked = False '开户银行
wbDN.txtZh.Locked = False '账号

For oo = 0 To 4
    wbDN.txtL(oo).Locked = False
Next

wbDN.frmJE.Visible = True
    If wbDN.optYz.Value = True Then
        wbDN.lblYZ.Caption = wbDN.txtKhmc.Text
        wbDN.txtAdr1.Text = wbDN.txtXmAdr.Text
    ElseIf wbDN.optWy.Value = True Then
        wbDN.lblWy.Caption = wbDN.txtKhmc.Text
    ElseIf wbDN.lblQT(1).Value = True Then
        wbDN.lblQT(1).Caption = wbDN.txtKhmc.Text
    ElseIf wbDN.lblQT(2).Value = True Then
        wbDN.lblQT(2).Caption = wbDN.txtKhmc.Text
    ElseIf wbDN.lblQT(3).Value = True Then
        wbDN.lblQT(3).Caption = wbDN.txtKhmc.Text
    ElseIf wbDN.lblQT(4).Value = True Then
        wbDN.lblQT(4).Caption = wbDN.txtKhmc.Text
    ElseIf wbDN.lblQT(5).Value = True Then
        wbDN.lblQT(5).Caption = wbDN.txtKhmc.Text
    End If
wbDN.comQy.Text = mod1.Qy
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
    frmHyxz.Height = 855
    frmHyxz.Width = 3525
    frmHyxz.Visible = False
    wbDN.Enabled = True
    Cancel = True
    frmHyxz.Visible = False
    If wbDN.txtKhdm.Text = "" Then
        wbDN.txtKhmc.SetFocus
    End If
End If

End Sub
