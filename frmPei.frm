VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmPei 
   BackColor       =   &H00C0FFC0&
   Caption         =   "培训资料"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdJS 
      Caption         =   "编辑"
      Height          =   285
      Left            =   900
      TabIndex        =   51
      Top             =   4110
      Width           =   495
   End
   Begin MSComCtl2.DTPicker dtpF 
      Height          =   255
      Left            =   1530
      TabIndex        =   46
      Top             =   1590
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "yy-MM-dd H:mm"
      Format          =   116064259
      CurrentDate     =   40805
   End
   Begin VB.Frame frmJs 
      BackColor       =   &H00C0FFC0&
      Caption         =   "讲师列表"
      Height          =   3825
      Left            =   4890
      TabIndex        =   42
      Top             =   5370
      Width           =   4515
      Begin VB.CommandButton cmdTDel 
         Caption         =   "删除"
         Height          =   315
         Left            =   3600
         TabIndex        =   50
         Top             =   3420
         Width           =   705
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgTN 
         Height          =   1275
         Left            =   2970
         TabIndex        =   49
         Top             =   1140
         Visible         =   0   'False
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   2249
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdJadd 
         Caption         =   "添加"
         Height          =   315
         Left            =   1440
         TabIndex        =   45
         Top             =   3420
         Width           =   675
      End
      Begin VB.TextBox txtJN 
         Height          =   285
         Left            =   120
         TabIndex        =   44
         Top             =   3420
         Width           =   1245
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgJs 
         Height          =   3075
         Left            =   30
         TabIndex        =   43
         Top             =   240
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   5424
         _Version        =   393216
         BackColor       =   12648384
         Rows            =   30
         FixedCols       =   0
         BackColorFixed  =   16777152
         BackColorBkg    =   12648384
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgT 
         Height          =   3075
         Left            =   2220
         TabIndex        =   48
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   5424
         _Version        =   393216
         BackColor       =   12648384
         Rows            =   30
         FixedCols       =   0
         BackColorFixed  =   16777152
         BackColorBkg    =   12648384
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H00C0FFC0&
      Caption         =   "修改"
      Height          =   765
      Left            =   11460
      Picture         =   "frmPei.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "修改"
      Top             =   8310
      Width           =   675
   End
   Begin VB.CommandButton cmdRdel 
      Caption         =   "人员删除"
      Height          =   315
      Left            =   4410
      TabIndex        =   40
      Top             =   8550
      Width           =   1185
   End
   Begin VB.CommandButton cmdBB1 
      Caption         =   "课程统计表"
      Height          =   315
      Left            =   9570
      TabIndex        =   39
      Top             =   7950
      Width           =   1245
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgRN 
      Height          =   1185
      Left            =   5550
      TabIndex        =   37
      Top             =   3360
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   2090
      _Version        =   393216
      Rows            =   100
      Cols            =   4
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.CommandButton cmdRen 
      Caption         =   "人员添加"
      Height          =   345
      Left            =   4440
      TabIndex        =   36
      Top             =   8160
      Width           =   1155
   End
   Begin VB.Frame frmPx 
      BackColor       =   &H00C0FFC0&
      Caption         =   "人员添加(双击列表中的名字)"
      Height          =   4425
      Left            =   4710
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   4635
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   4020
         Width           =   2085
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgRen 
         Height          =   3615
         Left            =   120
         TabIndex        =   34
         Top             =   300
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   6376
         _Version        =   393216
         BackColor       =   12648384
         Rows            =   30
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   16777152
         BackColorBkg    =   12648384
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   525
      Left            =   9660
      TabIndex        =   32
      Top             =   8460
      Visible         =   0   'False
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   926
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtXH 
      Height          =   315
      Left            =   6090
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   330
      Width           =   2955
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFC0&
      Caption         =   "添加"
      Height          =   765
      Left            =   12180
      Picture         =   "frmPei.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8310
      Width           =   705
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   4950
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5100
      Top             =   5580
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "作废"
      Enabled         =   0   'False
      Height          =   765
      Left            =   13620
      Picture         =   "frmPei.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8310
      Width           =   645
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   765
      Left            =   14280
      Picture         =   "frmPei.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8310
      Width           =   585
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "提交"
      Height          =   765
      Left            =   12900
      Picture         =   "frmPei.frx":09D8
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8310
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   7815
      Left            =   9390
      TabIndex        =   24
      Top             =   0
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   13785
      _Version        =   393216
      BackColor       =   12648384
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtFy 
      Height          =   315
      Left            =   6090
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   960
      Width           =   2955
   End
   Begin VB.TextBox txtMyd 
      Height          =   315
      Left            =   6090
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Text3"
      Top             =   1605
      Width           =   2955
   End
   Begin VB.TextBox txtPxpg 
      Height          =   315
      Left            =   6090
      TabIndex        =   21
      Text            =   "Text4"
      Top             =   2235
      Width           =   2955
   End
   Begin VB.TextBox txtZtime 
      Height          =   315
      Left            =   6090
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "Text5"
      Top             =   2865
      Width           =   2955
   End
   Begin VB.TextBox txtBz 
      Height          =   4725
      Left            =   6090
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Text            =   "frmPei.frx":1042
      Top             =   4080
      Width           =   2955
   End
   Begin VB.TextBox txtTeacher 
      Height          =   315
      Left            =   1530
      TabIndex        =   18
      Text            =   "Text7"
      Top             =   4080
      Width           =   2955
   End
   Begin VB.TextBox txtZbdw 
      Height          =   315
      Left            =   1530
      TabIndex        =   17
      Text            =   "Text6"
      Top             =   3455
      Width           =   2955
   End
   Begin VB.TextBox txtPxt 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   1530
      TabIndex        =   16
      Text            =   "Text5"
      Top             =   2175
      Width           =   2955
   End
   Begin VB.TextBox txtAdr 
      Height          =   315
      Left            =   1530
      TabIndex        =   15
      Text            =   "Text4"
      Top             =   2835
      Width           =   2955
   End
   Begin VB.TextBox txtMc 
      Height          =   315
      Left            =   1530
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   955
      Width           =   2955
   End
   Begin VB.TextBox txtLb 
      Height          =   315
      Left            =   1530
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   330
      Width           =   2955
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgA 
      Height          =   3825
      Left            =   120
      TabIndex        =   7
      Top             =   4710
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6747
      _Version        =   393216
      BackColor       =   12648384
      ForeColor       =   0
      Rows            =   30
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgAn 
      Height          =   4125
      Left            =   2160
      TabIndex        =   38
      Top             =   5940
      Visible         =   0   'False
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   7276
      _Version        =   393216
      BackColor       =   12648384
      Rows            =   30
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin MSComCtl2.DTPicker dtpL 
      Height          =   255
      Left            =   3150
      TabIndex        =   47
      Top             =   1590
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   450
      _Version        =   393216
      CustomFormat    =   "yy-MM-dd H:mm"
      Format          =   96665603
      CurrentDate     =   40805
   End
   Begin VB.Label lblPid 
      Caption         =   "lblPid"
      Height          =   255
      Left            =   6600
      TabIndex        =   31
      Top             =   3540
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "序号"
      Height          =   315
      Left            =   4980
      TabIndex        =   29
      Top             =   420
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "备注"
      Height          =   255
      Left            =   4980
      TabIndex        =   12
      Top             =   4140
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "总时数"
      Height          =   255
      Left            =   4980
      TabIndex        =   11
      Top             =   2925
      Width           =   1005
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "培训评估"
      Height          =   255
      Left            =   4980
      TabIndex        =   10
      Top             =   2295
      Width           =   1425
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "培训满意度"
      Height          =   255
      Left            =   4980
      TabIndex        =   9
      Top             =   1665
      Width           =   1005
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "培训费用"
      Height          =   255
      Left            =   4980
      TabIndex        =   8
      Top             =   1020
      Width           =   1425
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "讲师"
      Height          =   255
      Left            =   420
      TabIndex        =   6
      Top             =   4140
      Width           =   1665
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "主办单位"
      Height          =   255
      Left            =   420
      TabIndex        =   5
      Top             =   3515
      Width           =   1665
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "课程时数"
      Height          =   255
      Left            =   420
      TabIndex        =   4
      ToolTipText     =   "小时"
      Top             =   2265
      Width           =   1665
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "培训地点"
      Height          =   255
      Left            =   420
      TabIndex        =   3
      Top             =   2895
      Width           =   1665
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "培训时间"
      Height          =   255
      Left            =   420
      TabIndex        =   2
      Top             =   1640
      Width           =   1665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "课程名称"
      Height          =   255
      Left            =   420
      TabIndex        =   1
      Top             =   1015
      Width           =   1665
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "培训类别"
      Height          =   255
      Left            =   420
      TabIndex        =   0
      Top             =   390
      Width           =   1665
   End
End
Attribute VB_Name = "frmPei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer '1保存
Dim KF As String
Private Sub cmdAdd_Click()
Call Qing
cmdRen.Visible = False
cmdRdel.Visible = False
cmdSave.Enabled = True
'dtPT.Visible = True
End Sub

Private Sub cmdBack_Click()
Me.Visible = False

End Sub

Private Sub cmdBB1_Click()
Dim tt As String
Dim Ra
tt = "select * from peib1 order by pid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Call frmPeiBB1.Bound(Ra)
frmPeiBB1.Show
frmPeiBB1.ZOrder 0
End Sub

Private Sub cmdDel_Click()
Dim ii As Integer
Dim tt As String
On Error Resume Next
If Val(lblPid.Caption) = 0 Then Exit Sub
ii = MsgBox("是否删除此记录？", vbQuestion + vbYesNo, "请确认")
If ii = vbNo Then
    Exit Sub
End If
timZm = 5 '删除
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "培训资料"
    mod1.cmd.Parameters("@NBLX") = "删除"
    mod1.cmd.Parameters("@bh") = lblPid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""

    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = ""
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If

End Sub

Private Sub cmdJadd_Click()
Dim tt As String
tt = "select jname from peijs where jname='" & txtJN.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.BOF = False Then
    Set mod1.HTP = Nothing
    MsgBox ("讲师名字有重复！")
    Exit Sub
End If
Set mod1.HTP = Nothing

tt = "insert into peiJs (jname) values ('" & txtJN.Text & "')"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = Nothing
Call Me.dtgJsFF

End Sub

Private Sub cmdJS_Click()
Call Me.dtgJsFF
frmJs.Visible = True
End Sub

Private Sub cmdMod_Click()
If mod1.DName = "吴之禺" Or mod1.DName = "马晓聪" Or mod1.DName = "陈珊珊" And Val(lblPid.Caption) > 0 Then
    cmdSave.Enabled = True
    cmdDel.Enabled = True
    cmdRen.Visible = True
    cmdRdel.Visible = True
    cmdJS.Visible = True
End If
End Sub

Private Sub cmdRdel_Click()
Dim Uid As String
Dim tt As String
Dim ii As Integer
Dim Name As String
Dim liD As Long
dtgAn.Row = dtgA.Row
dtgAn.Col = 3: liD = Val(dtgAn.Text)
If liD = 0 Then Exit Sub

On Error Resume Next

dtgAn.Row = dtgA.Row
dtgAn.Col = 2
Uid = dtgAn.Text
dtgAn.Col = 1
Name = dtgAn.Text

If Left(Uid, 2) <> "HM" Then Exit Sub


timZm = 3 '人员添加
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "培训资料"
    mod1.cmd.Parameters("@NBLX") = "人员删除"
    mod1.cmd.Parameters("@bh") = lblPid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Name
    mod1.cmd.Parameters("@mt2") = Uid
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = liD
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If
cmdSave.Enabled = False
End Sub

Private Sub cmdRen_Click()
frmPx.Visible = True
Call Me.dtgRenFF
End Sub

Private Sub cmdSave_Click()
Dim ii As Integer
Dim tt As String
On Error Resume Next

If txtXH.Text = "" Then
    MsgBox "请输入序号!"
    Exit Sub
End If
cmdRen.Visible = True: frmJs.Visible = False
timZm = 1 '保存
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "培训资料"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = lblPid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtLb.Text
    mod1.cmd.Parameters("@mt2") = txtMc.Text
    mod1.cmd.Parameters("@mt3") = txtAdr.Text
    mod1.cmd.Parameters("@mt4") = txtZbdw.Text
    mod1.cmd.Parameters("@mt5") = txtTeacher.Text

    mod1.cmd.Parameters("@mt7") = txtPxpg.Text
    mod1.cmd.Parameters("@mlt1") = txtBz.Text
    mod1.cmd.Parameters("@mm1") = Val(txtPxt.Text)
    mod1.cmd.Parameters("@mm2") = Val(txtXH.Text)
    mod1.cmd.Parameters("@mm3") = Val(txtFy.Text)
    mod1.cmd.Parameters("@mm4") = Val(txtZtime.Text)
    mod1.cmd.Parameters("@mm5") = Val(lblPid.Caption)
    mod1.cmd.Parameters("@mm6") = Val(txtMyd.Text)
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = dtpF.Value
    mod1.cmd.Parameters("@md2") = dtpL.Value
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If
    
cmdSave.Enabled = False

End Sub

Private Sub cmdTDel_Click()
Dim Jid As Long
Dim tt As String
Dim RC
Dim Lc As Long
Jid = Val(dtgT.Text)
If Jid = 0 Then Exit Sub


tt = "delete from peijt where jid=" & Jid & ";" & _
    "select jid,jname from peiJDetail where pid=" & Val(lblPid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = mod1.HTP.NextRecordset
On Error Resume Next
RC = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Lc = UBound(RC, 2) + 1
Call Me.dtgTFF
dtgT.Rows = Lc + 10
txtTeacher.Text = ""
For oo = 1 To Lc
    txtTeacher.Text = txtTeacher.Text & " " & RC(1, oo - 1)
    dtgT.Row = oo
    dtgT.Col = 0: dtgT.Text = RC(0, oo - 1)
    dtgT.Col = 1: dtgT.Text = RC(1, oo - 1)
Next
Call cmdSave_Click
End Sub

Private Sub dtgA_DblClick()
Dim KQ As Boolean
Dim liD As Long
dtgAn.Row = dtgA.Row
dtgAn.Col = 3: liD = Val(dtgAn.Text)
If liD = 0 Then Exit Sub
dtgAn.Col = 1: KQ = dtgAn.Text
If KQ = "True" Then
    KQ = False
Else
    KQ = True
End If
On Error Resume Next

timZm = 3 '出勤否
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "培训资料"
    mod1.cmd.Parameters("@NBLX") = "出勤否"
    mod1.cmd.Parameters("@bh") = lblPid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = liD
    mod1.cmd.Parameters("@mm10") = Val(txtPxt.Text) '课程时数
    mod1.cmd.Parameters("@mb1") = KQ
    mod1.cmd.Parameters("@md1") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If
cmdSave.Enabled = False
End Sub


Private Sub dtgA_KeyDown(KeyCode As Integer, Shift As Integer)
Dim liD As Long

If KeyCode = 13 Then


dtgAn.Row = dtgA.Row
dtgAn.Col = 3: liD = Val(dtgAn.Text)
If liD = 0 Then Exit Sub
KF = InputBox("请输入考分:")
On Error GoTo MxcCss
On Error Resume Next

timZm = 6 '修改考分
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "培训资料"
    mod1.cmd.Parameters("@NBLX") = "修改考分"
    mod1.cmd.Parameters("@bh") = lblPid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = liD
    mod1.cmd.Parameters("@mm2") = Val(KF)
    mod1.cmd.Parameters("@mb1") = Null
    mod1.cmd.Parameters("@md1") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If
cmdSave.Enabled = False
End If
Exit Sub
MxcCss:

End Sub


Private Sub dtgBr_Click()
dtgN.Row = dtgBr.Row
If dtgN.Row = 0 Then Exit Sub
dtgN.Col = 4
lblPid.Caption = dtgN.Text
End Sub

Private Sub dtgBr_DblClick()
dtgN.Row = dtgBr.Row
If dtgN.Row = 0 Then Exit Sub
dtgN.Col = 4
If Val(dtgN.Text) = 0 Then Exit Sub
Call MXBound(Val(dtgN.Text))
End Sub


Private Sub dtgJs_DblClick()
Dim tt As String
Dim RC
Dim oo As Long
tt = "insert into peiJT (pid,jid) values (" & Val(lblPid.Caption) & "," & Val(dtgJs.Text) & ");" & _
    "select jid,jname from peiJDetail where pid=" & Val(lblPid.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Lc = UBound(RC, 2) + 1
Call Me.dtgTFF
dtgT.Rows = Lc + 10
txtTeacher.Text = ""
For oo = 1 To Lc
    txtTeacher.Text = txtTeacher.Text & " " & RC(1, oo - 1)
    dtgT.Row = oo
    dtgT.Col = 0: dtgT.Text = RC(0, oo - 1)
    dtgT.Col = 1: dtgT.Text = RC(1, oo - 1)
Next
Call cmdSave_Click
End Sub

Private Sub dtgRen_DblClick()
Dim Fy As Single
Dim Uid As String
Dim tt As String
Dim ii As Integer
Dim Name As String
dtgRN.Row = dtgRen.Row
dtgRN.Col = 0
Uid = dtgRN.Text
dtgRN.Col = 1
Name = dtgRN.Text

If Left(Uid, 2) <> "HM" Then Exit Sub
On Error Resume Next
tt = InputBox("请输入培训满意度！")
Fy = Val(tt)

timZm = 2 '人员添加
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "培训资料"
    mod1.cmd.Parameters("@NBLX") = "人员添加"
    mod1.cmd.Parameters("@bh") = lblPid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Name
    mod1.cmd.Parameters("@mt2") = Uid
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Fy '满意度
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If
cmdSave.Enabled = False
End Sub

Private Sub dtpF_Change()
On Error Resume Next
txtPxt.Text = Round(DateDiff("h", dtpF.Value, dtpL.Value), 1)
End Sub

Private Sub dtpL_Change()
On Error Resume Next
txtPxt.Text = Round(DateDiff("h", dtpF.Value, dtpL.Value), 1)
End Sub

Private Sub dtPT_CloseUp()
txtPxTime.Text = dtPT.Value
End Sub

Private Sub Form_Click()
frmPx.Visible = False
frmJs.Visible = False
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
frmRen.Left = 0
frmRen.Top = 0

End Sub

Public Sub Qing()
txtXH.Text = ""
txtLb.Text = ""
txtMc.Text = ""

txtAdr.Text = ""
txtPxt.Text = ""
txtZbdw.Text = ""
txtTeacher.Text = ""
txtFy.Text = ""
txtMyd.Text = ""
txtPxpg.Text = ""
txtZtime.Text = ""
txtBz.Text = ""

Call dtgAFF
Call dtgRenFF

'dtPT.Value = mod1.DQda
'dtPT.Visible = False
lblPid.Caption = ""
cmdSave.Enabled = False
cmdRen.Visible = False
cmdRdel.Visible = False
cmdDel.Enabled = False
txtFy.Locked = False
frmJs.Visible = False
txtTeacher.Locked = True
dtpF.Value = Date
dtpL.Value = Date
cmdJS.Visible = False
cmdRen.Visible = False
cmdRdel.Visible = False
End Sub


Public Sub dtgBRFF()
dtgBr.Clear
dtgBr.Cols = 5
dtgBr.Rows = 50
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "序号": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "培训类别": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "课程名称": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "讲师": dtgBr.CellFontBold = True
dtgBr.ColWidth(4) = 0
dtgBr.ColWidth(0) = 660
dtgBr.ColWidth(2) = 2745
dtgN.Clear
dtgN.Cols = 5
dtgN.Rows = 50
dtgN.Row = 0
dtgN.Col = 0: dtgN.Text = "序号": dtgN.CellFontBold = True
dtgN.Col = 1: dtgN.Text = "培训类别": dtgN.CellFontBold = True
dtgN.Col = 2: dtgN.Text = "课程名称": dtgN.CellFontBold = True
dtgN.Col = 3: dtgN.Text = "讲师": dtgN.CellFontBold = True
dtgN.ColWidth(4) = 0
dtgN.ColWidth(0) = 660
dtgN.ColWidth(2) = 2745
End Sub

Public Sub JLBound()
Dim tt As String
Dim oo As Integer
Dim Ra
Dim La As Long
tt = "select xh,lb,mc,teacher,pid from peixun order by pid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgBr.Visible = False
Call dtgBRFF
dtgBr.Rows = La + 50: dtgN.Rows = dtgBr.Rows
For oo = 1 To La
    dtgBr.Row = oo
    dtgBr.Col = 0: dtgBr.Text = Ra(0, oo - 1)
    dtgBr.Col = 1: dtgBr.Text = Ra(1, oo - 1)
    dtgBr.Col = 2: dtgBr.Text = Ra(2, oo - 1)
    dtgBr.Col = 3: dtgBr.Text = Ra(3, oo - 1)
    dtgBr.Col = 4: dtgBr.Text = Ra(4, oo - 1)
    dtgN.Row = oo
    dtgN.Col = 0: dtgN.Text = Ra(0, oo - 1)
    dtgN.Col = 1: dtgN.Text = Ra(1, oo - 1)
    dtgN.Col = 2: dtgN.Text = Ra(2, oo - 1)
    dtgN.Col = 3: dtgN.Text = Ra(3, oo - 1)
    dtgN.Col = 4: dtgN.Text = Ra(4, oo - 1)
Next
dtgBr.Visible = True
End Sub

Private Sub timQuit_Timer()
Dim Rz
Dim Lz As Integer
Dim Rb
Dim Lb As Integer
Dim RD
Dim Ld As Integer
Dim tt As String
On Error Resume Next
Dim ii As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '保存
    Call JLBound
ElseIf timZm = 2 Or timZm = 3 Or timZm = 1 Then
    tt = "select name,cf,uid,Lid,fy,myd,kf from peiren where pid=" & Val(lblPid.Caption) & " order by cf desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rb = mod1.HTP.GetRows
    Call Me.RenBound(Rb)
ElseIf timZm = 6 Then
    dtgA.Col = 6: dtgA.Text = KF
    dtgAn.Col = 6: dtgAn.Text = KF
End If
timQuit.Enabled = False
End Sub


Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim Bid As Long
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mm2,mt2,mt1,mt3,mt4 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    mod1.Ti = 5
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    timWait.Enabled = False
    If timZm = 1 Then
        txtMyd.Text = mod1.WP.Fields("mm2").Value
        txtZtime.Text = mod1.WP.Fields("mm3").Value
    ElseIf timZm = 2 Or timZm = 3 Then
        txtMyd.Text = mod1.WP.Fields("mm1").Value
        txtZtime.Text = mod1.WP.Fields("mm2").Value
    ElseIf timZm = 5 Then '删除
        Call Me.Qing
        Call Me.JLBound
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
'''''    If timZm = 1 Then
'''''        NiceButton1.Enabled = False
'''''    End If
    timWait.Enabled = False
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
'''''    If timZm = 1 Then
'''''        NiceButton1.Enabled = False
'''''    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub



Public Sub MXBound(Pid As Long)
Dim tt As String
Dim Ra
Dim Rb
Dim RC
Dim Lc
Dim oo As Integer
Call Qing

tt = "select xh,lb,mc,0,adr,pxt,zbdw,teacher,fy,myd,pxpg,ztime,bz,ft,lt from peixun where pid=" & Pid & ";" & _
    "select name,cf,uid,Lid,fy,myd,kf from peiren where pid=" & Pid & " order by cf desc;" & _
    "select jid,jname from peiJDetail where pid=" & Pid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Lc = UBound(RC, 2) + 1
txtXH.Text = Ra(0, 0)
txtLb.Text = Ra(1, 0)
txtMc.Text = Ra(2, 0)
'txtPxTime.Text = Ra(3, 0)
txtAdr.Text = Ra(4, 0)
txtPxt.Text = Me.strnum(Val(Ra(5, 0)))
txtZbdw.Text = Ra(6, 0)
'txtTeacher.Text = Ra(7, 0)
txtFy.Text = Ra(8, 0)
txtMyd.Text = Ra(9, 0)
txtPxpg.Text = Ra(10, 0)
txtZtime.Text = Ra(11, 0)
txtBz.Text = Ra(12, 0)
dtpF.Value = Ra(13, 0)
dtpL.Value = Ra(14, 0)
Call Me.RenBound(Rb)
lblPid.Caption = Pid

Call Me.dtgTFF
dtgT.Rows = Lc + 10
txtTeacher.Text = ""
For oo = 1 To Lc
    txtTeacher.Text = txtTeacher.Text & " " & RC(1, oo - 1)
    dtgT.Row = oo
    dtgT.Col = 0: dtgT.Text = RC(0, oo - 1)
    dtgT.Col = 1: dtgT.Text = RC(1, oo - 1)
Next
End Sub

Public Sub dtgRenFF()
dtgRen.Clear
dtgRen.Rows = 50
dtgRen.Cols = 3
dtgRen.Row = 0
dtgRen.Col = 0: dtgRen.Text = "工号": dtgRen.CellFontBold = True
dtgRen.Col = 1: dtgRen.Text = "姓名": dtgRen.CellFontBold = True
dtgRen.Col = 2: dtgRen.Text = "部门": dtgRen.CellFontBold = True

dtgRN.Clear
dtgRN.Rows = 50
dtgRN.Cols = 3

End Sub

Private Sub txtFy_Change()
On Error Resume Next
Dim oo As Integer
Dim jj As Single
Dim qq As Integer
For oo = 1 To 70
    dtgA.Row = oo
    dtgA.Col = 0
    If dtgA.Text = "" Then Exit For
Next
oo = oo - 1
jj = Round(Val(txtFy.Text) / oo, 2)
oo = oo + 1
For qq = 1 To oo - 1
    dtgA.Row = qq
    dtgA.Col = 4
    dtgA.Text = jj
Next
End Sub

Private Sub txtName_Change()
Dim Ra
Dim La As Integer
Dim tt As String
Dim oo As Integer
On Error Resume Next
If txtName.Text = "" Then Exit Sub
tt = "select userid,username,bm from worker where username like '%" & txtName.Text & "%'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Call Me.dtgRenFF
La = UBound(Ra, 2) + 1
For oo = 1 To La
    dtgRen.Row = oo
    dtgRen.Col = 0: dtgRen.Text = Ra(0, oo - 1)
    dtgRen.Col = 1: dtgRen.Text = Ra(1, oo - 1)
    dtgRen.Col = 2: dtgRen.Text = Ra(2, oo - 1)
    dtgRN.Row = oo
    dtgRN.Col = 0: dtgRN.Text = Ra(0, oo - 1)
    dtgRN.Col = 1: dtgRN.Text = Ra(1, oo - 1)
    dtgRN.Col = 2: dtgRN.Text = Ra(2, oo - 1)
Next
End Sub



Public Sub dtgAFF()
dtgA.Clear: dtgA.Row = 0
dtgA.Col = 0: dtgA.Text = "学员": dtgA.CellFontBold = True
dtgA.Col = 1: dtgA.Text = "出勤": dtgA.CellFontBold = True

dtgA.Col = 4: dtgA.Text = "培训费用": dtgA.CellFontBold = True
dtgA.Col = 5: dtgA.Text = "满意度": dtgA.CellFontBold = True
dtgA.Col = 6: dtgA.Text = "考分": dtgA.CellFontBold = True
dtgA.ColWidth(1) = 800
dtgA.ColWidth(2) = 0
dtgA.ColWidth(3) = 0
dtgA.ColWidth(4) = 900
dtgA.ColWidth(5) = 800
dtgA.ColWidth(6) = 800
dtgAn.Clear
End Sub

Public Sub RenBound(Rb)
On Error Resume Next
Dim Lb As Integer
Dim oo As Integer
Call Me.dtgAFF
Lb = UBound(Rb, 2) + 1
dtgA.Rows = Lb + 30
For oo = 1 To Lb
    dtgA.Row = oo
    dtgA.Col = 0: dtgA.Text = Rb(0, oo - 1): dtgA.CellForeColor = &H80000008
    dtgA.Col = 1: dtgA.Text = Rb(1, oo - 1): dtgA.CellForeColor = &H80000008
    dtgA.Col = 2: dtgA.Text = Rb(2, oo - 1): dtgA.CellForeColor = &H80000008
    dtgA.Col = 3: dtgA.Text = Rb(3, oo - 1): dtgA.CellForeColor = &H80000008
    dtgA.Col = 4: dtgA.Text = Rb(4, oo - 1): dtgA.CellForeColor = &H80000008
    dtgA.Col = 5: dtgA.Text = Rb(5, oo - 1): dtgA.CellForeColor = &H80000008
    dtgA.Col = 6: dtgA.Text = Rb(6, oo - 1): dtgA.CellForeColor = &H80000008
    dtgAn.Row = oo
    dtgAn.Col = 0: dtgAn.Text = Rb(0, oo - 1)
    dtgAn.Col = 1: dtgAn.Text = Rb(1, oo - 1)
    dtgAn.Col = 2: dtgAn.Text = Rb(2, oo - 1)
    dtgAn.Col = 3: dtgAn.Text = Rb(3, oo - 1)
    dtgAn.Col = 4: dtgAn.Text = Rb(4, oo - 1)
    dtgAn.Col = 5: dtgAn.Text = Rb(5, oo - 1)
    dtgAn.Col = 6: dtgAn.Text = Rb(6, oo - 1)
    dtgA.Col = 1
    If dtgA.Text = "False" Then
        dtgA.Text = "缺勤"
        dtgA.Col = 0: dtgA.CellForeColor = &HFF&
        dtgA.Col = 1: dtgA.CellForeColor = &HFF&
        dtgA.Col = 2: dtgA.CellForeColor = &HFF&
        dtgA.Col = 3: dtgA.CellForeColor = &HFF&
        dtgA.Col = 4: dtgA.CellForeColor = &HFF&
        dtgA.Col = 5: dtgA.CellForeColor = &HFF&
        dtgA.Col = 6: dtgA.CellForeColor = &HFF&
    Else
        dtgA.Text = "出勤"
    End If
Next

End Sub

Public Sub dtgJsFF()
Dim tt As String
Dim Ra
Dim La As Long
Dim oo As Long
dtgJs.Clear
dtgJs.Cols = 2
dtgJs.Col = 0: dtgJs.Row = 0: dtgJs.Text = "编号": dtgJs.CellFontBold = True
dtgJs.Col = 1: dtgJs.Row = 0: dtgJs.Text = "讲师姓名": dtgJs.CellFontBold = True
dtgJs.ColWidth(0) = 0
dtgJs.ColWidth(1) = 2000
tt = "select jid,jname from peijs order by jid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgJs.Rows = La + 30
For oo = 1 To La
    dtgJs.Row = oo
    dtgJs.Col = 0: dtgJs.Text = Ra(0, oo - 1)
    dtgJs.Col = 1: dtgJs.Text = Ra(1, oo - 1)
Next
End Sub
Public Sub dtgTFF()
Dim tt As String
Dim Ra
Dim La As Long
Dim oo As Long
dtgT.Clear
dtgT.Cols = 2
dtgT.Col = 0: dtgT.Row = 0: dtgT.Text = "编号": dtgT.CellFontBold = True
dtgT.Col = 1: dtgT.Row = 0: dtgT.Text = "讲师姓名": dtgT.CellFontBold = True
dtgT.ColWidth(0) = 0
dtgT.ColWidth(1) = 2000

End Sub
Function strnum(i As Single) As String
If Abs(i) < 1 And i <> 0 Then
If i > 0 Then
strnum = "0" & Trim(i)
Else
'''strnum = "-0" & Trim(Abs(i))
strnum = Trim(Abs(i))
End If
End If
strnum = i
End Function
