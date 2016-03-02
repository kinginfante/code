VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcXB 
   Caption         =   "销售报表"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   Begin VB.ComboBox comG 
      Height          =   300
      ItemData        =   "FmxcXB.frx":0000
      Left            =   780
      List            =   "FmxcXB.frx":0010
      TabIndex        =   45
      Text            =   "上海豪曼制冷空调服务有限公司"
      Top             =   8790
      Width           =   3375
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   420
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1950
      Top             =   0
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   285
      Left            =   11730
      TabIndex        =   42
      Top             =   8850
      Width           =   945
   End
   Begin VB.Frame frmMod 
      Caption         =   "财务到帐"
      Height          =   8565
      Left            =   9810
      TabIndex        =   7
      Top             =   90
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton cmdSave 
         Caption         =   "提交"
         Height          =   585
         Left            =   4230
         Picture         =   "FmxcXB.frx":0080
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   7800
         Width           =   675
      End
      Begin VB.TextBox txthtCD 
         Height          =   300
         Left            =   2040
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   7260
         Width           =   2775
      End
      Begin VB.TextBox txtWCF 
         Height          =   300
         Left            =   2040
         TabIndex        =   40
         Text            =   "Text1"
         Top             =   6780
         Width           =   2775
      End
      Begin VB.TextBox txtFHRQ 
         Height          =   300
         Left            =   2040
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   6270
         Width           =   2505
      End
      Begin VB.TextBox txtYJTQ 
         Height          =   300
         Left            =   2040
         TabIndex        =   37
         Text            =   "Text1"
         Top             =   5790
         Width           =   2775
      End
      Begin VB.TextBox txtFPQSD 
         Height          =   300
         Left            =   2040
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   5280
         Width           =   2775
      End
      Begin VB.TextBox txtKPCY 
         Height          =   300
         Left            =   2040
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   4800
         Width           =   2775
      End
      Begin VB.TextBox txtKPJE 
         Height          =   300
         Left            =   2040
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   4290
         Width           =   2775
      End
      Begin VB.TextBox txtKPRQ 
         Height          =   315
         Left            =   2040
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   3780
         Width           =   2505
      End
      Begin VB.TextBox txtKPSRQ 
         Height          =   315
         Left            =   2040
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   3300
         Width           =   2505
      End
      Begin VB.TextBox txtWBTZ 
         Height          =   300
         Left            =   2040
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   2790
         Width           =   2775
      End
      Begin VB.TextBox txtOJE 
         Height          =   300
         Left            =   2040
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   2310
         Width           =   2775
      End
      Begin VB.TextBox txtNje 
         Height          =   300
         Left            =   2040
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1830
         Width           =   2775
      End
      Begin VB.TextBox txtWKDJE 
         Height          =   300
         Left            =   2040
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtKDJE 
         Height          =   300
         Left            =   2040
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtkdRq 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   330
         Width           =   2505
      End
      Begin MSComCtl2.DTPicker dtpkdRq 
         Height          =   315
         Left            =   2040
         TabIndex        =   24
         Top             =   330
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Format          =   98959361
         CurrentDate     =   41403
      End
      Begin MSComCtl2.DTPicker dtpKPSRQ 
         Height          =   315
         Left            =   2040
         TabIndex        =   32
         Top             =   3300
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Format          =   98959361
         CurrentDate     =   41403
      End
      Begin MSComCtl2.DTPicker dtpKPRQ 
         Height          =   315
         Left            =   2040
         TabIndex        =   33
         Top             =   3780
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Format          =   98959361
         CurrentDate     =   41403
      End
      Begin MSComCtl2.DTPicker dtpFHRQ 
         Height          =   315
         Left            =   2040
         TabIndex        =   39
         Top             =   6270
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Format          =   98959361
         CurrentDate     =   41403
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H008080FF&
         Height          =   465
         Left            =   180
         Top             =   7590
         Width           =   4965
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "合同存档"
         Height          =   225
         Left            =   450
         TabIndex        =   22
         Top             =   7290
         Width           =   1125
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "是否完成"
         Height          =   255
         Left            =   180
         TabIndex        =   21
         Top             =   6810
         Width           =   1395
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "发货时间"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   6306
         Width           =   1395
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "佣金提取情况"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   5813
         Width           =   1395
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "发票签收单"
         Height          =   255
         Left            =   180
         TabIndex        =   18
         Top             =   5320
         Width           =   1395
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "开单开票差异"
         Height          =   255
         Left            =   180
         TabIndex        =   17
         Top             =   4827
         Width           =   1395
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "开票金额"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   4334
         Width           =   1395
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "开票日期"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   3841
         Width           =   1395
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "开票申请日期"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   3348
         Width           =   1395
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "维保进场通知"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   2855
         Width           =   1395
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "合同未结算金额"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   2362
         Width           =   1395
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "开单未结算金额"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   1869
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "未开单金额"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   1376
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "开单金额"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   883
         Width           =   1395
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "开单日期"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   390
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制"
      Height          =   255
      Left            =   10320
      TabIndex        =   6
      Top             =   8850
      Width           =   1155
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgXN 
      Height          =   405
      Left            =   12570
      TabIndex        =   5
      Top             =   8880
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   714
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "查询"
      Height          =   315
      Left            =   9180
      TabIndex        =   4
      Top             =   8850
      Width           =   975
   End
   Begin VB.TextBox txtZ 
      Height          =   270
      Left            =   6900
      TabIndex        =   3
      Top             =   8820
      Width           =   2085
   End
   Begin VB.ComboBox comLX 
      Height          =   300
      ItemData        =   "FmxcXB.frx":06EA
      Left            =   5370
      List            =   "FmxcXB.frx":06FA
      TabIndex        =   2
      Text            =   "年度"
      Top             =   8820
      Width           =   1365
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgX 
      Height          =   8685
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15319
      _Version        =   393216
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label17 
      Caption         =   "帐套"
      Height          =   285
      Left            =   90
      TabIndex        =   44
      Top             =   8820
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "查询方式"
      Height          =   315
      Left            =   4350
      TabIndex        =   1
      Top             =   8820
      Width           =   885
   End
End
Attribute VB_Name = "FmxcXB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Htbh As String
Dim Fid As Long
Dim NewF As Integer
Dim Hid As Long
Dim timZm As Integer
Dim ComanyId As Integer
Dim NYear As Integer

Private Sub cmdC_Click()
Dim tt As String
Dim Ra
Dim La As Long
    If comG.Text = "上海豪曼制冷空调服务有限公司" Then
        companyId = 1
    ElseIf comG.Text = "上海鼎力制冷空调设备有限公司" Then
        companyId = 2
    ElseIf comG.Text = "上海杰升商贸有限公司" Then
        companyId = 3
    ElseIf comG.Text = "广州杰狮机电设备有限公司" Then
        companyId = 4
    End If
Select Case comLX.Text
Case "年度"
    NYear = Val(txtZ.Text)
    tt = "select htbh,ddrq,khmc,xywy+'('+ywy+')',bm,htze,rq,yingfje,kdrq,kdje,wkdje,hxrq,hx,Nje,Oje,wbtz,kpsrq,kprq,kpje,kpcy,fpqsd,yjtq,fhrq,wcf,htcd,fid,newF,hid,kdfh" & _
        " from htx1 where year(ddrq)=" & txtZ.Text & " and companyId=" & companyId & " order by htbh,rq"
Case "合同编号"
    tt = "select htbh,ddrq,khmc,xywy+'('+ywy+')',bm,htze,rq,yingfje,kdrq,kdje,wkdje,hxrq,hx,Nje,Oje,wbtz,kpsrq,kprq,kpje,kpcy,fpqsd,yjtq,fhrq,wcf,fid,newF,hid,kdfh" & _
        " from htx1 where htbh='" & txtZ.Text & "' and companyId=" & companyId & " order by htbh,rq"
Case "合同金额"
    tt = "select htbh,ddrq,khmc,xywy+'('+ywy+')',bm,htze,rq,yingfje,kdrq,kdje,wkdje,hxrq,hx,Nje,Oje,wbtz,kpsrq,kprq,kpje,kpcy,fpqsd,yjtq,fhrq,wcf,fid,newF,hid,kdfh" & _
        " from htx1 where htze=" & Val(txtZ.Text) & " and companyId=" & companyId & " order by htbh,rq"
Case "月份"
    If NYear = 0 Then NYear = Year(mod1.DQda)
    tt = "select htbh,ddrq,khmc,xywy+'('+ywy+')',bm,htze,rq,yingfje,kdrq,kdje,wkdje,hxrq,hx,Nje,Oje,wbtz,kpsrq,kprq,kpje,kpcy,fpqsd,yjtq,fhrq,wcf,htcd,fid,newF,hid,kdfh" & _
        " from htx1 where year(ddrq)=" & NYear & " and month(ddrq)=" & Val(txtZ.Text) & " and companyId=" & companyId & " order by htbh,rq"
End Select

Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Call Me.Bound(Ra, La)

End Sub

Private Sub cmdCopy_Click()
dtgX.FixedRows = 0
dtgX.FixedCols = 0

dtgX.Col = 0
dtgX.Row = 0
dtgX.ColSel = 24
dtgX.RowSel = dtgX.Rows - 50
Clipboard.Clear
Clipboard.SetText dtgX.Clip
dtgX.FixedRows = 1
dtgX.FixedCols = 1
End Sub

Private Sub cmdMod_Click()
If mod1.DName <> "徐瑛" Then Exit Sub
frmMod.Visible = True

End Sub

Private Sub cmdSave_Click()
If Fid = 0 Then Exit Sub

timZm = 1 '保存
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "销售报表"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = Fid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtWBTZ.Text '维保进场通知
    mod1.cmd.Parameters("@mt2") = Val(txtKPCY.Text)  '开单开票差异
    mod1.cmd.Parameters("@mt3") = Val(txtYJTQ.Text)  '佣金提取情况
    mod1.cmd.Parameters("@mt4") = Val(txtWCF.Text)  '是否完成
    mod1.cmd.Parameters("@mt5") = Val(txthtCD.Text)  '合同存档
    mod1.cmd.Parameters("@mt6") = Val(txtFPQSD.Text)  '发票签收单
    
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtKDJE.Text) '开单金额
    mod1.cmd.Parameters("@mm2") = Val(txtWKDJE.Text)  '未开单金额
    mod1.cmd.Parameters("@mm3") = Val(txtNje.Text) '开单未结算金额
    mod1.cmd.Parameters("@mm4") = Val(txtOJE.Text) '合同未结算金额
    mod1.cmd.Parameters("@mm5") = Val(txtKPJE.Text) '开票金额
    mod1.cmd.Parameters("@mm6") = 0
    
    mod1.cmd.Parameters("@mb1") = 0
    If txtkdRq.Text = "" Then '开单日期
        mod1.cmd.Parameters("@md1") = Null
    Else
        mod1.cmd.Parameters("@md1") = txtkdRq.Text
    End If
    If txtKPSRQ.Text = "" Then '开票申请日期
        mod1.cmd.Parameters("@md2") = Null
    Else
        mod1.cmd.Parameters("@md2") = txtKPSRQ.Text
    End If
    If txtKPRQ.Text = "" Then '开票日期
        mod1.cmd.Parameters("@md3") = Null
    Else
        mod1.cmd.Parameters("@md3") = txtKPRQ.Text
    End If
    If txtFHRQ.Text = "" Then '发货时间
        mod1.cmd.Parameters("@md4") = Null
    Else
        mod1.cmd.Parameters("@md4") = txtFHRQ.Text
    End If
    
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 1 Then '保存
            cmdSave.Enabled = False
        End If
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True

        
    End If

    
Set mod1.cmd = Nothing


cmdSave.Enabled = False


End Sub

Private Sub dtgX_Click()
dtgXN.Row = dtgX.Row
dtgXN.Col = 0: Htbh = dtgXN.Text
dtgXN.Col = 25: Fid = Val(dtgXN.Text)
dtgXN.Col = 26: NewF = Val(dtgXN.Text)
dtgXN.Col = 27: Hid = Val(dtgXN.Text)

    dtgXN.Col = 8: txtkdRq.Text = dtgXN.Text
    dtgXN.Col = 9: txtKDJE.Text = dtgXN.Text
    dtgXN.Col = 10: txtWKDJE.Text = dtgXN.Text
    dtgXN.Col = 13: txtOJE.Text = dtgXN.Text
    dtgXN.Col = 14: txtNje.Text = dtgXN.Text
    dtgXN.Col = 15: txtWBTZ.Text = dtgXN.Text
    dtgXN.Col = 16: txtKPSRQ.Text = dtgXN.Text
    dtgXN.Col = 17: txtKPRQ.Text = dtgXN.Text
    dtgXN.Col = 18: txtKPJE.Text = dtgXN.Text
    dtgXN.Col = 19: txtKPCY.Text = dtgXN.Text
    dtgXN.Col = 20: txtFPQSD.Text = dtgXN.Text
    dtgXN.Col = 21: txtYJTQ.Text = dtgXN.Text
    dtgXN.Col = 22: txtFHRQ.Text = dtgXN.Text
    dtgXN.Col = 23: txtWCF.Text = dtgXN.Text
    dtgXN.Col = 24: txthtCD.Text = dtgXN.Text
End Sub

Private Sub dtgX_DblClick()
Dim tt As String
Dim xZ As String


Dim Bid As Long
Dim ZL As String
Dim Ra
'Dim Lid As String
On Error Resume Next



frmWait.Visible = True
frmWait.ZOrder 0
frmWait.Refresh



If dtgX.Col > 0 Then
    
    Exit Sub
End If
If NewF = 0 Then
    If xZ = "C. 维保合同" Or xZ = "D. 维修合同" Then
    'mod1.comJZ = False
    wbHTP.Visible = False
    Call modHt.wbQing
    
    
    tt = "Select * from htping where hid=" & Hid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Call modHt.wbBound
    
    
    '打开材料表
    tt = "Select * from htSale where htbh='" & wbHTP.txtHtbh.Text & "'"
    wbMx.adoRGF.Recordset.Close
    wbMx.adoRGF.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set wbMx.dtgSale.DataSource = wbMx.adoRGF
    wbMx.lblChg.Caption = wbHTP.txtClcb1.Text
    
    '打开应收款表
    tt = "Select * from htping1 where htBh='" & wbHTP.lblHid.Caption & "' order by rq"
    frmFuK.adoHpt.Recordset.Close
    frmFuK.adoHpt.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set wbMx.dtgFk.DataSource = frmFuK.adoHpt
    
    '打开佣金表
    tt = "Select * from Yongjin where htBh='" & wbHTP.txtHtbh.Text & "' order by yId"
    frmYj.adoYj.Recordset.Close
    frmYj.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set frmYj.dtgYJ.DataSource = frmYj.adoYj
    
    ''打开出工信息表(如果为评审阶段则不显示）
    'If wbHTP.optZ.Value = True Or wbHTP.optW.Value = True Then
    '    tt = "Select max(gzb.rq),max(gzb.wxWorker),sum(workXX.wTime),max(bhid)" & _
    '    "max(htbh) from gzb cross join workXX where gzb.bhid=workXX.bhid and gzb.htBh='" & _
    '    wbHTP.txtHtbh.Text & "' group by gzb.bhid"
    '    form2Htp.adoGzb.Recordset.Close
    '    form2Htp.adoGzb.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '    Set wbMx.dtgGzb.DataSource = form2Htp.adoGzb
    'End If
    wbHTP.txtYj1.Visible = False
    wbHTP.txtYj2.Visible = False
    wbHTP.txtLr1.Visible = False
    wbHTP.txtLr2.Visible = False
    wbHTP.lblTcBe.Visible = False
    wbHTP.txtTcBe.Visible = False
    wbHTP.UpDa.Visible = False
    wbHTP.lblYj.Visible = False
    wbHTP.lblLr.Visible = False
    wbHTP.lblTC.Visible = False
    wbHTP.Visible = True
    Exit Sub
    End If
    
    
    
    
    
    
    
    
    
    
    '购销合同
    
    form2Htp.Visible = True
    mod1.workTt = ""
    mod1.workTt = "Select * from htPing where hid=" & Hid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open mod1.workTt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    form2Htp.lblHtxz.Caption = ""
    
    Call modHt.htQing
    Call modHt.htBound '绑定合同评审单字段
    

    
    
    '打开收款表
    
    
    tt = "Select * from htPing1 where htBh='" & form2Htp.lblHid.Caption & "' order by rq"
    frmFuK.adoHpt.Recordset.Close
    frmFuK.adoHpt.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    
    
    Set frmFuK.dtgFk.DataSource = frmFuK.adoHpt
    
    'ft = "Select * from yiFk Where htBh='" & frmFuK.adoHpt.Recordset.Fields("htBh").Value & _
    '"' and yingRQ='" & frmFuK.adoHpt.Recordset.Fields("rq").Value & "' order by yiRq"
    'frmFuK.adoYf.Recordset.Close
    'frmFuK.adoYf.Recordset.Open ft, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'Set frmFuK.dtgYf.DataSource = frmFuK.adoYf
    
    '打开产品表
    tt = ""
    tt = "Select * from htSale Where htBh='" & form2Htp.txtHtbh.Text & "'"
    form2Htp.adoSale.Recordset.Close
    form2Htp.adoSale.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set form2Htp.dtgSale.DataSource = form2Htp.adoSale
    Set form2Htp.dtgYJ.DataSource = form2Htp.adoSale
    Set form2Htp.dtgZj.DataSource = form2Htp.adoSale
    
    ''打开“取自库存表”
    'tt = "Select * from kcJa where htBh='" & form2Htp.txtHtbh.Text & "'"
    'form2Htp.adoKu.Recordset.Close
    'form2Htp.adoKu.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'Set form2Htp.dtgKu.DataSource = form2Htp.adoKu
    
    ''打开采购表
    'ft = "Select * from CG Where htbh='" & form2Htp.txtHtbh.Text & "' and khmc<>'库存'"
    'frmAdo.adoTmp.Recordset.Close
    'frmAdo.adoTmp.Recordset.Open ft, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'Set form2Htp.dtgCG.DataSource = frmAdo.adoTmp
    
    '打开佣金表
    tt = "Select * from Yongjin where htBh='" & form2Htp.txtHtbh.Text & "' order by yId"
    frmYj.adoYj.Recordset.Close
    frmYj.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set frmYj.dtgYJ.DataSource = frmYj.adoYj
    
    
    
    
    form2Htp.tabHt.TabEnabled(1) = True
    form2Htp.tabHt.TabEnabled(2) = True
    'End If
    
    
    
    
    
    
    
    form2Htp.tabHt.Tab = 0
    htBrowG.MousePointer = 0
    
    
        '佣金、利润2、提成不显示
        form2Htp.txtYj1.Visible = False
        form2Htp.txtYj2.Visible = False
        form2Htp.txtLr1.Visible = False
        form2Htp.txtLr2.Visible = False
        'form2Htp.txtTc1.Visible = False
        'form2Htp.txtTc2.Visible = False
        form2Htp.lblYj.Visible = False
        form2Htp.lblLr2.Visible = False
        'form2Htp.lblTc.Visible = False
ElseIf NewF = 1 Then
        Call modHt.NewQing
        Call modHt.NewLocked
        Call modHt.NewBound(Hid)
'            '设置流程按钮
'        If (frmWbNew.lblHtxz = "维保" And frmWbNew.txtHtze > 50000) Or Val(frmWbNew.txtHtze.Text) > 10000 Then
'            Call modHt.HtLcBut(63)
'        Else
'            Call modHt.HtLcBut(62)
'        End If
        frmWbNew.Visible = True
ElseIf NewF = 2 Then
        Call modNewHT.NewMQing
        Call modNewHT.NewLocked
        Call modNewHT.NewMBound(Hid)
        FMXC.lblMQM(0).Visible = True
        FMXC.lblMTm(0).Visible = True
        FMXC.cmdMQm(0).Visible = True
        
ElseIf NewF = 3 Or NewF = 5 Or NewF = 7 Then
        Call modNewHT.NewMQing
        Call modNewHT.NewLocked
        Call modNewHT.NewB(Hid)
        FMXC.lblMQM(0).Visible = True
        FMXC.lblMTm(0).Visible = True
        FMXC.cmdMQm(0).Visible = True
ElseIf NewF = 6 Or NewF = 8 Then
    Call FmxcNew.Bound(Hid)
    FmxcNew.Show
    FmxcNew.ZOrder 0
End If
End Sub


Private Sub dtpFHRQ_CloseUp()
txtFHRQ.Text = DateSerial(Year(dtpFHRQ.Value), Month(dtpFHRQ.Value), Day(dtpFHRQ.Value))
End Sub


Private Sub dtpkdRq_CloseUp()
txtkdRq.Text = DateSerial(Year(dtpkdRq.Value), Month(dtpkdRq.Value), Day(dtpkdRq.Value))
End Sub

Private Sub dtpKPRQ_CloseUp()
txtKPRQ.Text = DateSerial(Year(dtpKPRQ.Value), Month(dtpKPRQ.Value), Day(dtpKPRQ.Value))
End Sub


Private Sub dtpKPSRQ_CloseUp()
txtKPSRQ.Text = DateSerial(Year(dtpKPSRQ.Value), Month(dtpKPSRQ.Value), Day(dtpKPSRQ.Value))
End Sub


Private Sub Form_Click()
frmMod.Visible = False
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth + 500
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
dtgX.Left = 0
dtgX.Top = 0
Call dtgXFF
txtZ.Text = Year(mod1.DQda)

dtpKPSRQ.Value = mod1.DQda
dtpKPRQ.Value = mod1.DQda
dtpkdRq.Value = mod1.DQda
dtpFHRQ.Value = mod1.DQda

End Sub

Public Sub dtgXFF()
dtgX.Cols = 28

dtgX.Clear
dtgX.Row = 0
dtgX.Col = 0: dtgX.Text = "单据编号": dtgX.CellFontBold = True
dtgX.Col = 1: dtgX.Text = "单据日期": dtgX.CellFontBold = True
dtgX.Col = 2: dtgX.Text = "客户": dtgX.CellFontBold = True
dtgX.Col = 3: dtgX.Text = "业务员": dtgX.CellFontBold = True
dtgX.Col = 4: dtgX.Text = "部门": dtgX.CellFontBold = True
dtgX.Col = 5: dtgX.Text = "总金额": dtgX.CellFontBold = True
dtgX.Col = 6: dtgX.Text = "应收日期": dtgX.CellFontBold = True
dtgX.Col = 7: dtgX.Text = "应收金额": dtgX.CellFontBold = True
dtgX.Col = 8: dtgX.Text = "开单日期": dtgX.CellFontBold = True
dtgX.Col = 9: dtgX.Text = "开单金额": dtgX.CellFontBold = True
dtgX.Col = 10: dtgX.Text = "未开单金额": dtgX.CellFontBold = True
dtgX.Col = 11: dtgX.Text = "结算日期": dtgX.CellFontBold = True
dtgX.Col = 12: dtgX.Text = "结算金额": dtgX.CellFontBold = True
dtgX.Col = 13: dtgX.Text = "开单未结算金额": dtgX.CellFontBold = True
dtgX.Col = 14: dtgX.Text = "合同未结算金额": dtgX.CellFontBold = True
dtgX.Col = 15: dtgX.Text = "维保进场通知": dtgX.CellFontBold = True
dtgX.Col = 16: dtgX.Text = "开票申请日期": dtgX.CellFontBold = True
dtgX.Col = 17: dtgX.Text = "开票日期": dtgX.CellFontBold = True
dtgX.Col = 18: dtgX.Text = "开票金额": dtgX.CellFontBold = True
dtgX.Col = 19: dtgX.Text = "开单开票差异": dtgX.CellFontBold = True
dtgX.Col = 20: dtgX.Text = "发票签收单": dtgX.CellFontBold = True
dtgX.Col = 21: dtgX.Text = "佣金提取情况": dtgX.CellFontBold = True
dtgX.Col = 22: dtgX.Text = "发货时间": dtgX.CellFontBold = True
dtgX.Col = 23: dtgX.Text = "是否完成": dtgX.CellFontBold = True
dtgX.Col = 24: dtgX.Text = "合同存档": dtgX.CellFontBold = True
dtgX.Col = 25: dtgX.Text = "fid": dtgX.CellFontBold = True
dtgX.Col = 26: dtgX.Text = "NewF": dtgX.CellFontBold = True
dtgX.Col = 27: dtgX.Text = "hid": dtgX.CellFontBold = True
dtgXN.Cols = 28

dtgXN.Clear
dtgXN.Row = 0
dtgXN.Col = 0: dtgXN.Text = "单据编号": dtgXN.CellFontBold = True
dtgXN.Col = 1: dtgXN.Text = "单据日期": dtgXN.CellFontBold = True
dtgXN.Col = 2: dtgXN.Text = "客户": dtgXN.CellFontBold = True
dtgXN.Col = 3: dtgXN.Text = "业务员": dtgXN.CellFontBold = True
dtgXN.Col = 4: dtgXN.Text = "部门": dtgXN.CellFontBold = True
dtgXN.Col = 5: dtgXN.Text = "总金额": dtgXN.CellFontBold = True
dtgXN.Col = 6: dtgXN.Text = "应收日期": dtgXN.CellFontBold = True
dtgXN.Col = 7: dtgXN.Text = "应收金额": dtgXN.CellFontBold = True
dtgXN.Col = 8: dtgXN.Text = "开单日期": dtgXN.CellFontBold = True
dtgXN.Col = 9: dtgXN.Text = "开单金额": dtgXN.CellFontBold = True
dtgXN.Col = 10: dtgXN.Text = "未开单金额": dtgXN.CellFontBold = True
dtgXN.Col = 11: dtgXN.Text = "结算日期": dtgXN.CellFontBold = True
dtgXN.Col = 12: dtgXN.Text = "结算金额": dtgXN.CellFontBold = True
dtgXN.Col = 13: dtgXN.Text = "开单未结算金额": dtgXN.CellFontBold = True
dtgXN.Col = 14: dtgXN.Text = "合同未结算金额": dtgXN.CellFontBold = True
dtgXN.Col = 15: dtgXN.Text = "维保进场通知": dtgXN.CellFontBold = True
dtgXN.Col = 16: dtgXN.Text = "开票申请日期": dtgXN.CellFontBold = True
dtgXN.Col = 17: dtgXN.Text = "开票日期": dtgXN.CellFontBold = True
dtgXN.Col = 18: dtgXN.Text = "开票金额": dtgXN.CellFontBold = True
dtgXN.Col = 19: dtgXN.Text = "开单开票差异": dtgXN.CellFontBold = True
dtgXN.Col = 20: dtgXN.Text = "发票签收单": dtgXN.CellFontBold = True
dtgXN.Col = 21: dtgXN.Text = "佣金提取情况": dtgXN.CellFontBold = True
dtgXN.Col = 22: dtgXN.Text = "发货时间": dtgXN.CellFontBold = True
dtgXN.Col = 23: dtgXN.Text = "是否完成": dtgXN.CellFontBold = True
dtgXN.Col = 24: dtgXN.Text = "合同存档": dtgXN.CellFontBold = True
dtgXN.Col = 25: dtgXN.Text = "fid": dtgXN.CellFontBold = True
dtgXN.Col = 26: dtgXN.Text = "NewF": dtgXN.CellFontBold = True
dtgXN.Col = 27: dtgXN.Text = "hid": dtgXN.CellFontBold = True
dtgX.ColWidth(25) = 0
dtgX.ColWidth(26) = 0
dtgX.ColWidth(27) = 0
dtgX.ColWidth(0) = 1900
dtgX.ColWidth(2) = 3600
dtgX.ColWidth(3) = 1650
Call Me.ModQing
End Sub

Public Sub Bound(Ra, La As Long)
Dim oo As Long
Dim ii As Long
Dim THTBH As String
On Error Resume Next
dtgX.Visible = False
Call Me.dtgXFF
Me.dtgX.Rows = La + 100
Me.dtgXN.Rows = La + 100
For oo = 1 To La
    dtgX.Row = oo
    dtgXN.Row = oo
    For ii = 0 To 28
        dtgX.Col = ii
        dtgXN.Col = ii
        dtgX.Text = Ra(ii, oo - 1)
        dtgXN.Text = Ra(ii, oo - 1)
        If ii = 1 Or ii = 6 Or ii = 8 Or ii = 11 Or ii = 16 Or ii = 17 Then
            dtgX.Text = DateSerial(Year(Ra(ii, oo - 1)), Month(Ra(ii, oo - 1)), Day(Ra(ii, oo - 1)))
        End If
        
    Next
        If oo = 1 Then THTBH = Ra(0, 0)
        If oo > 1 Then
            If Ra(0, oo - 1) = THTBH Then
                dtgX.Col = 5: dtgX.Text = ""
                dtgXN.Col = 5: dtgXN.Text = ""
            Else
                THTBH = Ra(0, oo - 1)
            End If
        End If
        If Ra(28, oo - 1) = "True" Then
            dtgX.Col = 6: dtgX.Text = "款到发货"
            dtgXN.Col = 6: dtgXN.Text = "款到发货"
        End If
Next
dtgX.Visible = True
End Sub

Public Sub ModQing()
txtkdRq.Text = ""
txtKDJE.Text = ""
txtWKDJE.Text = ""
txtOJE.Text = ""
txtNje.Text = ""
txtWBTZ.Text = ""
txtKPSRQ.Text = ""
txtKPRQ.Text = ""
txtKPJE.Text = ""
txtKPCY.Text = ""
txtFPQSD.Text = ""
txtYJTQ.Text = ""
txtFHRQ.Text = ""
txtWCF.Text = ""
txthtCD.Text = ""
End Sub

Private Sub Form_Resize()
   dtgX.Width = Me.Width - 200
End Sub

Private Sub timQuit_Timer()
Dim oo As Integer
Dim tt As String
Dim ii As Integer
Dim RC
Dim Lc As Integer
On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0

If timZm = 1 Then '保存"
    cmdSave.Enabled = True

    frmMod.Visible = False
    
    dtgX.Col = 8: dtgX.Text = txtkdRq.Text
    dtgX.Col = 9: dtgX.Text = txtKDJE.Text
    dtgX.Col = 10: dtgX.Text = txtWKDJE.Text
    dtgX.Col = 13: dtgX.Text = txtOJE.Text
    dtgX.Col = 14: dtgX.Text = txtNje.Text
    dtgX.Col = 15: dtgX.Text = txtWBTZ.Text
    dtgX.Col = 16: dtgX.Text = txtKPSRQ.Text
    dtgX.Col = 17: dtgX.Text = txtKPRQ.Text
    dtgX.Col = 18: dtgX.Text = txtKPJE.Text
    dtgX.Col = 19: dtgX.Text = txtKPCY.Text
    dtgX.Col = 20: dtgX.Text = txtFPQSD.Text
    dtgX.Col = 21: dtgX.Text = txtYJTQ.Text
    dtgX.Col = 22: dtgX.Text = txtFHRQ.Text
    dtgX.Col = 23: dtgX.Text = txtWCF.Text
    dtgX.Col = 24: dtgX.Text = txthtCD.Text
    dtgXN.Col = 8: dtgXN.Text = txtkdRq.Text
    dtgXN.Col = 9: dtgXN.Text = txtKDJE.Text
    dtgXN.Col = 10: dtgXN.Text = txtWKDJE.Text
    dtgXN.Col = 13: dtgXN.Text = txtOJE.Text
    dtgXN.Col = 14: dtgXN.Text = txtNje.Text
    dtgXN.Col = 15: dtgXN.Text = txtWBTZ.Text
    dtgXN.Col = 16: dtgXN.Text = txtKPSRQ.Text
    dtgXN.Col = 17: dtgXN.Text = txtKPRQ.Text
    dtgXN.Col = 18: dtgXN.Text = txtKPJE.Text
    dtgXN.Col = 19: dtgXN.Text = txtKPCY.Text
    dtgXN.Col = 20: dtgXN.Text = txtFPQSD.Text
    dtgXN.Col = 21: dtgXN.Text = txtYJTQ.Text
    dtgXN.Col = 22: dtgXN.Text = txtFHRQ.Text
    dtgXN.Col = 23: dtgXN.Text = txtWCF.Text
    dtgXN.Col = 24: dtgXN.Text = txthtCD.Text
    txtkdRq.Text = ""
    txtKDJE.Text = ""
    txtWKDJE.Text = ""
    txtOJE.Text = ""
    txtNje.Text = ""
    txtWBTZ.Text = ""
    txtKPSRQ.Text = ""
    txtKPRQ.Text = ""
    txtKPJE.Text = ""
    txtKPCY.Text = ""
    txtFPQSD.Text = ""
    txtFHRQ.Text = ""
    txtWCF.Text = ""
    txthtCD.Text = ""
End If
timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2,mt3 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then


    ElseIf timZm = 5 Then

    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    timWait.Enabled = False
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdSave.Enabled = False
    End If
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdSave.Enabled = False
    End If
    Exit Sub

End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


