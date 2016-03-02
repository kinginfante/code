VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWBXX 
   Caption         =   "人工费询价单"
   ClientHeight    =   9225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15210
   Begin VB.Frame frmTj 
      Caption         =   "商务部调价"
      Height          =   2565
      Left            =   1080
      TabIndex        =   49
      Top             =   5460
      Width           =   6915
      Begin VB.TextBox txtT6 
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   1650
         TabIndex        =   62
         Top             =   1680
         Width           =   1845
      End
      Begin VB.ComboBox txtT1 
         Height          =   300
         ItemData        =   "frmWBXX.frx":0000
         Left            =   1650
         List            =   "frmWBXX.frx":0025
         TabIndex        =   60
         Text            =   "Combo1"
         Top             =   390
         Width           =   1935
      End
      Begin VB.CommandButton cmdTJ 
         Caption         =   "提交"
         Height          =   285
         Left            =   5760
         TabIndex        =   59
         Top             =   2220
         Width           =   885
      End
      Begin VB.TextBox txtT5 
         Height          =   1695
         Left            =   3780
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Top             =   360
         Width           =   2925
      End
      Begin VB.TextBox txtT4 
         Height          =   315
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   2100
         Width           =   1905
      End
      Begin VB.TextBox txtT3 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1650
         TabIndex        =   54
         Top             =   1255
         Width           =   1875
      End
      Begin VB.TextBox txtT2 
         Height          =   315
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   785
         Width           =   1875
      End
      Begin VB.Label Label9 
         Caption         =   "调整金额(差旅)"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   1740
         Width           =   1275
      End
      Begin VB.Label Label8 
         Caption         =   "情况说明"
         Height          =   195
         Left            =   3780
         TabIndex        =   57
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "调整结果"
         Height          =   225
         Left            =   270
         TabIndex        =   55
         Top             =   2190
         Width           =   945
      End
      Begin VB.Label Label6 
         Caption         =   "调整金额(人工)"
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   120
         TabIndex        =   53
         Top             =   1350
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "  原基准价  (包含差旅费)"
         Height          =   405
         Left            =   90
         TabIndex        =   51
         Top             =   810
         Width           =   1155
      End
      Begin VB.Label Label2 
         Caption         =   "针对业务"
         Height          =   255
         Left            =   330
         TabIndex        =   50
         Top             =   420
         Width           =   795
      End
   End
   Begin VB.TextBox txt2 
      Height          =   285
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   8760
      Width           =   1665
   End
   Begin VB.TextBox txtLadr 
      Height          =   315
      Left            =   1050
      TabIndex        =   44
      Top             =   5550
      Width           =   5595
   End
   Begin VB.Frame frmHide 
      Caption         =   "frmHid"
      Height          =   1455
      Left            =   1680
      TabIndex        =   14
      Top             =   7590
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label lblHtbh 
         Caption         =   "lblHtbh"
         Height          =   195
         Left            =   3300
         TabIndex        =   43
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label lblHLC 
         Caption         =   "lblHLC"
         Height          =   345
         Left            =   2670
         TabIndex        =   42
         Top             =   1080
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblYwy 
         Caption         =   "lblYwy"
         Height          =   285
         Left            =   2490
         TabIndex        =   22
         Top             =   780
         Width           =   765
      End
      Begin VB.Label lblUid 
         Caption         =   "lblUid"
         Height          =   255
         Left            =   3750
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   1860
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   930
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblLcRen 
         Caption         =   "lblLcRen"
         Height          =   285
         Left            =   150
         TabIndex        =   18
         Top             =   510
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblNlb 
         Caption         =   "lblNlb"
         Height          =   225
         Left            =   1680
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   315
         Left            =   1080
         TabIndex        =   16
         Top             =   450
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label lblBid 
         Caption         =   "lblBid"
         Height          =   255
         Left            =   210
         TabIndex        =   15
         Top             =   150
         Visible         =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame frmAdd 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   7950
      TabIndex        =   38
      Top             =   5130
      Width           =   7155
      Begin VB.OptionButton opt2 
         Caption         =   "外包"
         Height          =   195
         Left            =   1410
         TabIndex        =   64
         Top             =   90
         Width           =   735
      End
      Begin VB.OptionButton opt1 
         Caption         =   "本公司人工"
         Height          =   180
         Left            =   60
         TabIndex        =   63
         Top             =   90
         Width           =   1245
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H008080FF&
         Caption         =   "删除"
         Height          =   285
         Left            =   6420
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEx 
         BackColor       =   &H00C0FFC0&
         Caption         =   "新添询价业务"
         Height          =   285
         Left            =   5130
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   30
         Width           =   1245
      End
      Begin VB.ComboBox comLx 
         ForeColor       =   &H00FF0000&
         Height          =   300
         ItemData        =   "frmWBXX.frx":00A9
         Left            =   3180
         List            =   "frmWBXX.frx":00D1
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   0
         Width           =   1785
      End
      Begin VB.Label Label23 
         Caption         =   "业务类别"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   2250
         TabIndex        =   41
         Top             =   60
         Width           =   915
      End
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   4710
      TabIndex        =   0
      Top             =   6150
      Visible         =   0   'False
      Width           =   6315
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1365
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   300
         Width           =   4965
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   2
         Top             =   870
         Width           =   675
      End
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   2505
      Left            =   6930
      TabIndex        =   37
      Top             =   5520
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4419
      _Version        =   393216
      BackColor       =   12648447
      ForeColor       =   8404992
      Rows            =   15
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   16761024
      ForeColorFixed  =   0
      BackColorBkg    =   12648447
      GridColorFixed  =   8404992
      GridColorUnpopulated=   8404992
      WordWrap        =   -1  'True
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton cmdHT 
      BackColor       =   &H00C0FFC0&
      Caption         =   "合同评审单"
      Height          =   435
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   210
      Width           =   1065
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   180
      Top             =   7170
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   660
      Top             =   5790
   End
   Begin VB.TextBox txtBz 
      Height          =   1935
      Left            =   1050
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   6060
      Width           =   5595
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   13380
      Picture         =   "frmWBXX.frx":014B
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "修改"
      Top             =   8820
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   13890
      Picture         =   "frmWBXX.frx":0455
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "保存"
      Top             =   8820
      Width           =   465
   End
   Begin VB.CommandButton cmdQm 
      Height          =   345
      Index           =   0
      Left            =   6930
      TabIndex        =   11
      Top             =   8430
      Width           =   945
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14790
      Picture         =   "frmWBXX.frx":0ABF
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "返回"
      Top             =   8790
      Width           =   465
   End
   Begin VB.CommandButton cmdQm 
      Height          =   345
      Index           =   1
      Left            =   8010
      TabIndex        =   9
      Top             =   8430
      Width           =   945
   End
   Begin VB.CommandButton cmdQm 
      Height          =   345
      Index           =   2
      Left            =   9090
      TabIndex        =   8
      Top             =   8430
      Width           =   945
   End
   Begin VB.TextBox comXmmc 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   210
      Width           =   5025
   End
   Begin VB.CommandButton cmdD 
      Enabled         =   0   'False
      Height          =   405
      Left            =   14280
      Picture         =   "frmWBXX.frx":0BC1
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8790
      Width           =   465
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   405
      Left            =   -60
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   714
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgNew 
      Height          =   4335
      Left            =   0
      TabIndex        =   36
      Top             =   1020
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   7646
      _Version        =   393216
      Rows            =   3
      Cols            =   5
      FixedCols       =   0
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label lblZl 
      Caption         =   "Label10"
      Height          =   315
      Left            =   10380
      TabIndex        =   65
      Top             =   270
      Width           =   1485
   End
   Begin VB.Label lbl2 
      Caption         =   "基准价格"
      Height          =   375
      Left            =   90
      TabIndex        =   47
      Top             =   8790
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "地址"
      Height          =   225
      Left            =   270
      TabIndex        =   45
      Top             =   5610
      Width           =   495
   End
   Begin VB.Label lblTX 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      TabIndex        =   35
      Top             =   8250
      Width           =   5475
   End
   Begin VB.Label lblBz 
      Caption         =   "备注"
      Height          =   225
      Left            =   270
      TabIndex        =   34
      Top             =   6180
      Width           =   495
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   6930
      TabIndex        =   33
      Top             =   8820
      Width           =   945
   End
   Begin VB.Label lblQM 
      Caption         =   "业务员"
      Height          =   225
      Index           =   0
      Left            =   6960
      TabIndex        =   32
      Top             =   8130
      Width           =   1005
   End
   Begin VB.Label lblBh 
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7770
      TabIndex        =   31
      Top             =   195
      Width           =   1725
   End
   Begin VB.Label Label5 
      Caption         =   "编号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7020
      TabIndex        =   30
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Label4 
      Caption         =   "项目名称"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   29
      Top             =   270
      Width           =   975
   End
   Begin VB.Label lblQM 
      Caption         =   "商务支持"
      Height          =   225
      Index           =   1
      Left            =   8040
      TabIndex        =   28
      Top             =   8130
      Width           =   1005
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   8010
      TabIndex        =   27
      Top             =   8820
      Width           =   945
   End
   Begin VB.Label lblQM 
      Caption         =   "业务员确认"
      Height          =   225
      Index           =   2
      Left            =   9120
      TabIndex        =   26
      Top             =   8130
      Width           =   1005
   End
   Begin VB.Label lblTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   2
      Left            =   9090
      TabIndex        =   25
      Top             =   8820
      Width           =   945
   End
End
Attribute VB_Name = "frmWBXX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer '(1业务添加 2业务删除 3新人工签字 5表单保存 6调整添加 8删除)
Dim Mid As Long
Dim LX As String

Private Sub cmdBack_Click()
Me.Visible = False
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub

Private Sub cmdD_Click()
Dim ii As Integer
Dim tt As String
On Error Resume Next
tt = "select htbh from htping where hid=" & Val(lblHtbh.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.Fields(0).Value <> "HMNEW" Then
    Exit Sub
End If
mod1.HTP.Close
Set mod1.HTP = Nothing
If lblYwy.Caption <> mod1.DName Then Exit Sub
ii = MsgBox("是否删除此询价单？", vbYesNo + vbQuestion, "Hello")
If ii = vbNo Then
    Exit Sub
End If
timZm = 8 '删除合同
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "删除"
    mod1.cmd.Parameters("@bh") = lblBid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(lblZl.Caption)
    mod1.cmd.Parameters("@mm1") = 0
    mod1.cmd.Parameters("@mm2") = Val(lblHtbh.Caption)
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

Private Sub cmdDel_Click()
Dim oo As Integer
Dim ii As Integer
Dim Mid As Long
On Error Resume Next
dtgN.Col = 5
Mid = Val(dtgN.Text)
If Mid = 0 Then Exit Sub
ii = MsgBox("是否确定删除此记录?", vbQuestion + vbYesNo, "您好")
If ii = vbNo Then Exit Sub


 '业务删除
    timZm = 2
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "业务删除"
    mod1.cmd.Parameters("@bh") = lblBid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "人工询价"
    mod1.cmd.Parameters("@mt2") = comLx.Text  '业务类型

    mod1.cmd.Parameters("@mm1") = Mid

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


    Set mod1.cmd = Nothing

End Sub

Private Sub cmdDing_Click()
Dim tt As String
On Error Resume Next
If OptT1.Value = False And optT2.Value = False Then
    Exit Sub
End If
If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
timZm = 3 '新人工签字
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "新人工签字"
    mod1.cmd.Parameters("@bh") = Val(lblBid.Caption)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = Trim(lblYwy.Caption)
    mod1.cmd.Parameters("@mt2") = Trim(lblUid.Caption)
    mod1.cmd.Parameters("@mt3") = Trim(comXmmc.Text)
    mod1.cmd.Parameters("@mt4") = Trim(lblHtbh.Caption)
    mod1.cmd.Parameters("@mt5") = lblZl.Caption
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = ""
    mod1.cmd.Parameters("@mt11") = ""
    mod1.cmd.Parameters("@mt12") = ""
    mod1.cmd.Parameters("@mt13") = ""
    mod1.cmd.Parameters("@mt14") = Trim(lblFwid.Caption)
    mod1.cmd.Parameters("@mt15") = ""
    mod1.cmd.Parameters("@mt16") = ""
    mod1.cmd.Parameters("@mt17") = ""
    mod1.cmd.Parameters("@mt18") = ""
    mod1.cmd.Parameters("@mt19") = mod1.Qy
    mod1.cmd.Parameters("@mt20") = lblQM(Val(lblLc.Caption) - 1).Caption
    mod1.cmd.Parameters("@mt21") = ""
    
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = txtQM.Text '评审建议
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(lblLc.Caption)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 0
    mod1.cmd.Parameters("@mm11") = 0
    mod1.cmd.Parameters("@mm12") = 0
    mod1.cmd.Parameters("@mm13") = 0
    mod1.cmd.Parameters("@mm14") = 0
    mod1.cmd.Parameters("@mm15") = 0
    mod1.cmd.Parameters("@mm16") = Val(txt2.Text) '基准价格
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    If OptT1.Value = True Then
        mod1.cmd.Parameters("@mb1") = 1 '同意
    Else
        mod1.cmd.Parameters("@mb1") = 0 '拒绝
    End If
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = Null
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        cmdDing.Enabled = False
    
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
End Sub

Private Sub cmdEx_Click()
Dim oo As Integer
Dim ii As Integer
Dim Mid As Long
On Error Resume Next
dtgN.Col = 0
If comLx.Text <> "主机维保" And comLx.Text <> "小机末端空调箱保养" Then
    MsgBox "此项业务功能还需完善,请在备注中列明,由商务部核价!"
    Exit Sub
End If
For oo = 1 To dtgN.Rows
    dtgN.Row = oo
    If Trim(dtgN.Text) = comLx.Text Then
        ii = MsgBox("此业务已经添加,是否对它进行编辑?", vbQuestion + vbYesNo + vbDefaultButton1, "询价")
        If ii = vbYes Then
            dtgN.Col = 5
            Mid = Val(dtgN.Text)
            If comLx.Text = "主机维保" Then
                Call frmWBXT.Qing
                Call frmWBXT.Bound(Mid)
                frmWBXT.Show: frmWBXT.ZOrder 0
            ElseIf comLx.Text = "小机末端空调箱保养" Then
                Call frmWBXT2.Qing
                Call frmWBXT2.Bound(Mid)
                frmWBXT2.Show: frmWBXT2.ZOrder 0
            End If

            Exit Sub
        Else
            Exit Sub
        End If
    End If
Next


 '业务添加
    timZm = 1
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "业务添加"
    mod1.cmd.Parameters("@bh") = lblBid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "人工询价"
    mod1.cmd.Parameters("@mt2") = comLx.Text  '业务类型
''''''    mod1.cmd.Parameters("@mt4") = ""
''''''    mod1.cmd.Parameters("@mt5") = ""
''''''    mod1.cmd.Parameters("@mt6") = ""
''''''    mod1.cmd.Parameters("@mt7") = ""
''''''    mod1.cmd.Parameters("@mt8") = ""
''''''    mod1.cmd.Parameters("@mt9") = ""
''''''    mod1.cmd.Parameters("@mt10") = ""
''''''    mod1.cmd.Parameters("@mt11") = ""
''''''    mod1.cmd.Parameters("@mt12") = ""
''''''    mod1.cmd.Parameters("@mt13") = ""
''''''    mod1.cmd.Parameters("@mt14") = ""
''''''    mod1.cmd.Parameters("@mt15") = ""
''''''    mod1.cmd.Parameters("@mt16") = ""
''''''    mod1.cmd.Parameters("@mt17") = ""
''''''    mod1.cmd.Parameters("@mt18") = ""
''''''    mod1.cmd.Parameters("@mt19") = ""
''''''    mod1.cmd.Parameters("@mt20") = ""
''''''    mod1.cmd.Parameters("@mt21") = ""
''''''    mod1.cmd.Parameters("@mt22") = ""
''''''    mod1.cmd.Parameters("@mt23") = ""
''''''    mod1.cmd.Parameters("@mt24") = ""
''''''    mod1.cmd.Parameters("@mt25") = ""
''''''    mod1.cmd.Parameters("@mlt1") = ""
''''''    mod1.cmd.Parameters("@mlt2") = ""
''''''    mod1.cmd.Parameters("@mlt3") = ""
''''''    mod1.cmd.Parameters("@mlt4") = ""
''''''    mod1.cmd.Parameters("@mlt5") = ""
''''''    mod1.cmd.Parameters("@mm1") = 0
''''''    mod1.cmd.Parameters("@mm2") = 0
''''''    mod1.cmd.Parameters("@mm3") = 0
''''''    mod1.cmd.Parameters("@mm4") = 0
''''''    mod1.cmd.Parameters("@mm5") = 0
''''''    mod1.cmd.Parameters("@mm6") = 0
''''''    mod1.cmd.Parameters("@mm7") = 0
''''''    mod1.cmd.Parameters("@mm8") = 0
''''''    mod1.cmd.Parameters("@mm9") = 0
''''''    mod1.cmd.Parameters("@mm10") = 0
''''''    mod1.cmd.Parameters("@mm11") = 0
''''''    mod1.cmd.Parameters("@mm12") = 0
''''''    mod1.cmd.Parameters("@mm13") = 0
''''''    mod1.cmd.Parameters("@mm14") = 0
''''''    mod1.cmd.Parameters("@mm15") = 0
''''''    mod1.cmd.Parameters("@mm16") = 0
''''''    mod1.cmd.Parameters("@mm17") = 0
''''''    mod1.cmd.Parameters("@mm18") = 0
''''''    mod1.cmd.Parameters("@mm19") = 0
''''''    mod1.cmd.Parameters("@mm20") = 0
''''''    mod1.cmd.Parameters("@mb1") = 0
''''''    mod1.cmd.Parameters("@mb2") = 0
''''''    mod1.cmd.Parameters("@mb3") = 0
''''''    mod1.cmd.Parameters("@mb4") = 0
''''''    mod1.cmd.Parameters("@mb5") = 0
''''''    mod1.cmd.Parameters("@md1") = Null
''''''    mod1.cmd.Parameters("@md2") = Null
''''''    mod1.cmd.Parameters("@md3") = Null
''''''    mod1.cmd.Parameters("@md4") = Null
''''''    mod1.cmd.Parameters("@md5") = Null
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


    Set mod1.cmd = Nothing

End Sub

Private Sub cmdHt_Click()
Dim Ra
Dim tt As String
tt = "select newf from htping where hid=" & Val(lblHtbh.Caption)
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
If Ra(0, 0) = 6 Then
    Call FmxcNew.Bound(Val(lblHtbh.Caption))
    FmxcNew.Show
    FmxcNew.ZOrder 0
    Me.Visible = False
    Exit Sub
End If
If mod1.DName = "张砚纯" Or mod1.Gzu > 0 Then
    Exit Sub
End If
If mod1.DName = "彭海翔" And lblYwy.Caption <> mod1.DName Then '彭海翔只能打开自己的合同
    MsgBox "哈哈！"
    MsgBox "你想干嘛？"
    Exit Sub
End If
mod1.BTZ = 6

If FMXC.Visible = True And Val(FMXC.lblMHid.Caption) = Val(lblHtbh.Caption) Then
    Me.Visible = False
    FMXC.Enabled = True
    FMXC.ZOrder 0
Else

        Call modNewHT.NewMQing
        
        Call modNewHT.NewB(Val(lblHtbh.Caption))
        If FMXC.Visible = True Then '如果打开成功,则隐藏自己.
            Me.Visible = False
            FMXC.ZOrder 0
        End If
End If
    FMXC.cmdMQm(0).Visible = True
    FMXC.lblMQM(0).Visible = True
    FMXC.lblMTm(0).Visible = True
    FMXC.ZOrder 0
End Sub

Private Sub cmdMod_Click()

If FMXC.txtXYwy.ToolTipText = mod1.DHid And Val(lblLc.Caption) < 3 Then
    cmdSave.Enabled = True
    
    If Val(lblLc.Caption) = 1 Then
        txtLadr.Locked = False
        txtBz.Locked = False
        frmAdd.Visible = True
        cmdD.Enabled = True
    End If

End If
If mod1.DName = "" Or lblLcRen.Caption = "贾锦红" Or mod1.DName = "马晓聪" Or mod1.DName = "杨燕" Then
        txtLadr.Locked = False
        txtBz.Locked = False


   cmdSave.Enabled = True
    txtBz.Locked = False
    frmTj.Visible = True
    frmAdd.Visible = True

End If

If mod1.DName = FMXC.txtXYwy.Text Or mod1.DName = "马晓聪" Then
    cmdDel.Enabled = True
    txtT3.Locked = False
End If
If mod1.DName = "马晓聪" Then '马晓聪可以修改成本，并将成本导入合同
    lblLc.Caption = 3
    frmAdd.Visible = True
    txtBz.Locked = False
    txtLadr.Locked = False
    frmTj.Visible = True
End If


End Sub

Private Sub cmdQm_Click(Index As Integer)
Dim hg As Single
Dim Ra
Dim ii As Integer: Dim oo As Integer
On Error Resume Next
cmdDing.Enabled = True
If Index + 1 <> lblLc.Caption Then '不能在不相干的位置上乱点
    Exit Sub
End If

If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If
If lblLcUid.Caption <> mod1.DHid Then
'''''    tt = "select xuid from htping where hid=" & Val(lblHtbh.Caption)
'''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''    Ra = mod1.HTP.GetRows
'''''    mod1.HTP.Close
'''''    Set mod1.HTP = Nothing
'''''    If Ra(0, 0) <> mod1.DHid Then
        MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
        Exit Sub
'''''    End If
End If

hg = 0
dtgN.Row = 1: dtgN.Col = 1
For oo = 1 To dtgN.Rows
    dtgN.Row = oo: dtgN.Col = 1
    hg = hg + Val(dtgN.Text)
    dtgN.Col = 2
    hg = hg + Val(dtgN.Text)
Next
If Val(txt2.Text) <> hg Then
    cmdSave.Enabled = True
    Exit Sub
End If

frmQm.Visible = True
If lblLc.Caption = 1 Then
    optT2.Enabled = False
    OptT1.Value = True
Else
    optT2.Enabled = True
    OptT1.Value = False
    optT2.Value = False
End If
End Sub

Private Sub cmdQm_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tt As String
On Error Resume Next
If Button = 2 And lblQM(Index).Caption = "业务员确认" And Val(lblLc.Caption) = 100 And FMXC.txtXYwy.Text = mod1.DName Then

    If Val(lblHLC.Caption) < 2 Then
        Me.frmQm.Visible = True
        Me.OptT1.Enabled = False
        Me.optT2.Enabled = True
        Me.optT2.Value = True
        Me.lblLc.Caption = 3
            optT2.Caption = "增补"
    Else
        optT2.Caption = "拒绝"

    End If

End If
End Sub


Private Sub cmdSave_Click()
Dim oo As Integer
Dim ii As Integer
On Error Resume Next


Call ji

 '表单保存
    timZm = 5
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "表单保存"
    mod1.cmd.Parameters("@bh") = lblBid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtLadr.Text '工作地址
    mod1.cmd.Parameters("@mlt1") = txtBz.Text '备注
    mod1.cmd.Parameters("@mm1") = Val(txt2.Text) '基准总价

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


    Set mod1.cmd = Nothing

End Sub

Private Sub cmdTj_Click()
Dim oo As Integer
Dim ii As Integer
Dim Mid As Long
On Error Resume Next

If txtT5.Text = "" Or txtT1.Text = "" Or txtT1.Text = "商务部调价" Then
    Exit Sub
End If

 '业务添加
    timZm = 6
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "询价单"
    mod1.cmd.Parameters("@NBLX") = "调整添加"
    mod1.cmd.Parameters("@bh") = lblBid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "人工询价"
    mod1.cmd.Parameters("@mt2") = txtT1.Text   '业务类型
    mod1.cmd.Parameters("@mm1") = Val(txtT3.Text) '调整价格人工
    mod1.cmd.Parameters("@mm2") = Val(txtT6.Text) '调整价格差旅
    mod1.cmd.Parameters("@mlt1") = "商务部调价:" & txtT5.Text
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


    Set mod1.cmd = Nothing
End Sub

Private Sub dtgNew_Click()
'frmWBXT2.Show
dtgN.Row = dtgNew.Row
If dtgN.Row = 0 Then Exit Sub
dtgN.Col = 0
LX = dtgN.Text
dtgN.Col = 5: Mid = Val(dtgN.Text)
dtgN.Col = 4

'If frmTj.Visible = False Then Exit Sub

dtgN.Col = 0
txtT1.Text = dtgN.Text
dtgN.Col = 1
txtT2.Text = Val(dtgN.Text)
dtgN.Col = 2
txtT2.Text = Val(txtT2.Text) + Val(dtgN.Text)
txtT4.Text = txtT2.Text
End Sub

Private Sub dtgNew_DblClick()
dtgN.Col = 4
If Left(dtgN.Text, 5) = "商务部调价" Then Exit Sub
dtgN.Col = 4
If Left(dtgN.Text, 5) = "商务部调价" Then Exit Sub
If LX = "主机维保" Then
    Call frmWBXT.Qing
    Call frmWBXT.Bound(Mid)
    frmWBXT.Show
    frmWBXT.ZOrder 0
ElseIf LX = "小机末端空调箱保养" Then
    Call frmWBXT2.Qing
    Call frmWBXT2.Bound(Mid)
    frmWBXT2.Show
    frmWBXT2.ZOrder 0
End If
End Sub

Private Sub Form_DblClick()
frmQm.Visible = False
frmTj.Visible = False
End Sub

Private Sub Form_Load()
Dim oo As Integer
Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
dtgNew.Cols = 7
dtgN.Cols = 7
frmQm.Left = 6900
frmQm.Top = 7470
frmTj.Left = 0
frmTj.Top = 5490
End Sub

Public Sub MXBound(Bid As Long)
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim Ra
Dim La
On Error GoTo EPPWbxx3
tt = "SELECT dbo.MLMX.mt2 as 业务内容, dbo.MLMX.mM1 as 基准金额, dbo.MLMX.mM2 as 交通差旅费, dbo.MLMX.mt5 as 承接人, dbo.MLMX.mLT1 as 备注, dbo.MLMX.mid,dbo.wbPxNew.xz" & _
     " FROM dbo.wbPxNew INNER JOIN dbo.MLMX ON dbo.wbPxNew.DX = dbo.MLMX.mt2 where dbo.MLMX.bid=" & Bid & " order by dbo.wbPxNew.Zid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.BOF = False Then
    Ra = mod1.HTP.GetRows
    La = UBound(Ra, 2) + 1
End If
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next

dtgNew.Clear: dtgN.Clear
Call Me.dtgNewFF
dtgNew.Rows = La + 20
dtgN.Rows = dtgNew.Rows: dtgN.Cols = dtgNew.Cols

For oo = 1 To La + 1
    dtgNew.Row = oo: dtgN.Row = oo
    For ii = 0 To 5
        dtgNew.Col = ii: dtgN.Col = ii
        dtgNew.Text = Ra(ii, oo - 1)
        dtgN.Text = Ra(ii, oo - 1)
        dtgN.Col = 4
        If Left(dtgN.Text, 5) = "商务部调价" Then
            dtgNew.Col = 0: dtgNew.CellForeColor = &H8000000D
            dtgNew.Col = 1: dtgNew.CellForeColor = &H8000000D
            dtgNew.Col = 4: dtgNew.CellForeColor = &H8000000D
        End If
    Next
Next


   

Exit Sub
EPPWbxx3:
MsgBox ("网络故障，请退出后再试！")
End
End Sub

Public Sub Qing()

comXmmc.Text = ""
lblBh.Caption = ""
cmdHT.ToolTipText = ""
txtBz.Text = ""
lblTX.Caption = ""
lblYwy.Caption = ""
lblUid.Caption = ""
lblFwid.Caption = ""
lblLc.Caption = ""
lblLcRen.Caption = ""
lblLcUid.Caption = ""
lblNlb.Caption = ""
lblBid.Caption = ""

txt2.Text = ""
txtQM.Text = ""
OptT1.Value = False
optT2.Value = False
OptT1.Enabled = True
optT2.Enabled = True

cmdQm(0).Caption = ""
lblTm(0).Caption = ""
cmdQm(1).Caption = ""
lblTm(1).Caption = ""
cmdQm(2).Caption = ""
lblTm(2).Caption = ""

Call dtgPFF
Call dtgNewFF
dtgNew.Rows = 30

frmAdd.Visible = False
lblHtbh.Caption = ""
txtBz.Locked = True
txtLadr.Text = ""
txtLadr.Locked = True
frmTj.Visible = False

txtT1.ToolTipText = ""
txtT2.Text = ""
txtT3.Text = ""
txtT4.Text = ""
txtT5.Text = ""
txtT6.Text = ""
opt1.Value = False
opt2.Value = False
comLx.Locked = True
End Sub
Public Sub Bound(Bid As Long)
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim Ra, Rb, RC, RD, RE
Dim La, Lb, Lc, Ld, Le
Dim EntC As Integer
On Error GoTo EPPwbxx2
mod1.BTZ = 36

tt = "declare @hid int;" & _
    "select @hid=cast(htbh as int) from xunjiaD where bid=" & Bid & ";" & _
    "select zl,xid,xmmc,bid,bianhao,hg,jhg,ywy,uid,lc,lcren,lcuid,fwid,nlb,bz,htbh,yfadr from XunJiaD where bid=" & Bid & ";" & _
    "select lc from htping where hid=@hid;" & _
    "SELECT dbo.MLMX.mt2 as 业务内容, dbo.MLMX.mM1 as 基准金额, dbo.MLMX.mM2 as 交通差旅费, dbo.MLMX.mt5 as 承接人, dbo.MLMX.mLT1 as 备注, dbo.MLMX.mid,dbo.wbPxNew.xz" & _
     " FROM dbo.wbPxNew INNER JOIN dbo.MLMX ON dbo.wbPxNew.DX = dbo.MLMX.mt2 where dbo.MLMX.bid=" & Bid & " order by dbo.wbPxNew.Zid;" & _
    "select * from QMRZ where btz=36 and qdbh='" & Bid & "' order by zid;" & _
    "select trq,ywy,zn,bz,tf from pizu where bh='" & Bid & "' and yid=43 order by pid desc"

 
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
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
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next
La = UBound(Ra, 2) + 1
Lc = UBound(RC, 2) + 1

Ld = UBound(RD, 2) + 1
Le = UBound(RE, 2)

lblZl.Caption = Ra(0, 0)
comXmmc.Tag = Ra(1, 0)
comXmmc.Text = Ra(2, 0)
lblBid.Caption = Ra(3, 0)
lblBh.Caption = Ra(4, 0)
txt2.Text = Ra(6, 0)
lblYwy.Caption = Ra(7, 0)
lblUid.Caption = Ra(8, 0)
lblLc.Caption = Ra(9, 0)
lblLcRen.Caption = Ra(10, 0)
lblLcUid.Caption = Ra(11, 0)
lblFwid.Caption = Ra(12, 0)
lblNlb.Caption = Ra(13, 0)
txtBz.Text = Ra(14, 0)
lblHtbh.Caption = Ra(15, 0)
txtLadr.Text = Ra(16, 0)

lblHLC.Caption = Rb(0, 0) '对应合同的流程

''''''''列表明细
dtgNew.Clear: dtgN.Clear
Call Me.dtgNewFF
dtgNew.Rows = La + 20
dtgN.Rows = dtgNew.Rows: dtgN.Cols = dtgNew.Cols

For oo = 1 To Lc + 1
    dtgNew.Row = oo: dtgN.Row = oo
    EntC = 0
    For ii = 0 To 5
        dtgNew.Col = ii: dtgN.Col = ii
        dtgNew.Text = RC(ii, oo - 1)
        dtgN.Text = RC(ii, oo - 1)
        dtgN.Col = 4
        EntC = Len(dtgNew.Text) - Len(Replace(dtgNew.Text, Chr(13), ""))
        If Left(dtgN.Text, 5) = "商务部调价" Then
            dtgNew.Col = 0: dtgNew.CellForeColor = &H8000000D
            dtgNew.Col = 1: dtgNew.CellForeColor = &H8000000D
            dtgNew.Col = 4: dtgNew.CellForeColor = &H8000000D
        End If
        If ii = 4 Then
            If Len(dtgNew.Text) > 30 Or EntC > 0 Then
                If UpInt(Len(dtgNew.Text) / 30) > EntC Then
                    dtgNew.RowHeight(oo) = dtgNew.RowHeight(oo) * (UpInt(Len(dtgNew.Text) / 30) + 2)
                Else
                    dtgNew.RowHeight(oo) = dtgNew.RowHeight(oo) * EntC
                End If
            End If
        End If
    Next
Next

'签字按钮
For oo = 0 To 2
cmdQm(oo).Caption = ""
lblTm(oo).Caption = ""
Next
 For oo = 0 To Ld - 1
    If RD(9, oo) = True Then
       cmdQm(oo).Caption = RD(1, oo)
       lblTm(oo).Caption = RD(4, oo)
    End If
   cmdQm(oo).Tag = RD(8, oo)
Next

dtgP.Rows = Le + 20
dtgP.Clear
For oo = 1 To Le + 1
    dtgP.Row = oo
    For ii = 0 To 5
        dtgP.Col = ii
        dtgP.Text = RE(ii, oo - 1)
        If ii = 3 Then
            If Len(RE(ii, oo - 1)) > 16 Then
                dtgP.RowHeight(oo) = UpInt(Len(RE(ii, oo - 1)) / 16) * dtgP.RowHeight(oo)
            End If
        End If
        If ii = 4 Then
            If dtgP.Text = "True" Then
                dtgP.Text = "同意"
            ElseIf dtgP.Text = "False" Then
                dtgP.Text = "驳回"
            End If

        End If
    Next
Next
For oo = 1 To Le + 1
    dtgP.Row = oo
    dtgP.Col = 4
            If dtgP.Text = "驳回" Then
                For ii = 0 To 4
                    dtgP.Col = ii
                    dtgP.CellForeColor = &HFF&
                Next
            End If
Next
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "日期": dtgP.Col = 1: dtgP.Text = "姓名": dtgP.Col = 2: dtgP.Text = "职能"
dtgP.Col = 3: dtgP.Text = "评审建议": dtgP.Col = 4: dtgP.Text = "通过否"

cmdSave.Enabled = False
opt1.Enabled = False: opt2.Enabled = False
txtT1.Text = ""
For oo = 10 To 0 Step -1
    txtT1.RemoveItem oo
Next
If lblZl.Caption = "人工" Then
    opt1.Enabled = True: opt1.Value = True
    txtT1.AddItem "主机维保"
    txtT1.AddItem "主机大修"
    txtT1.AddItem "溴化锂维保"
    txtT1.AddItem "恒温恒湿机维保"
    txtT1.AddItem "小机维修"
    txtT1.AddItem "小机末端空调箱保养"
    txtT1.AddItem "水泵保养"
    txtT1.AddItem "冷却塔保养"
    txtT1.AddItem "电机保养"

ElseIf lblZl.Caption = "分包" Then
    opt2.Enabled = True: opt2.Value = True
    txtT1.AddItem "外包"
    txtT1.AddItem "人员常驻"
    txtT1.AddItem "水处理"
End If

Exit Sub
EPPwbxx2:
MsgBox ("网络故障，请退出后重试！")
End
End Sub


Public Sub dtgNewFF()
dtgNew.Clear
dtgN.Clear
Dim oo As Integer
For oo = 1 To dtgNew.Rows - 1
    dtgNew.RowHeight(oo) = dtgNew.RowHeight(0)
Next
dtgNew.Row = 0: dtgNew.Col = 0: dtgNew.Text = "业务内容"
dtgNew.Col = 1: dtgNew.Text = "基准金额"
dtgNew.Col = 2: dtgNew.Text = "交通差旅费"
dtgNew.Col = 3: dtgNew.Text = " 承接人"
dtgNew.Col = 4: dtgNew.Text = " 备注"
dtgNew.ColWidth(0) = 2000
dtgNew.ColWidth(4) = 9690
dtgNew.ColWidth(5) = 0
dtgNew.ColWidth(6) = 0

End Sub

Private Sub opt1_Click()
Dim oo As Integer
On Error Resume Next
If lblZl.Caption = "分包" Then
    opt2.Value = True
    Exit Sub
End If
For oo = 20 To 0 Step -1
    comLx.RemoveItem oo
Next


If opt1.Value = True Then
    comLx.AddItem "主机维保"
    comLx.AddItem "主机大修"
    comLx.AddItem "溴化锂"
    comLx.AddItem "恒温恒湿机"
    comLx.AddItem "小机维修"
    comLx.AddItem "小机末端空调箱保养"
    comLx.AddItem "水泵"
    comLx.AddItem "冷却塔"
    comLx.AddItem "电机"
    comLx.AddItem "交通差旅"
    comLx.AddItem "人员常驻"
    comLx.Locked = False
End If
End Sub

Private Sub opt2_Click()
Dim oo As Integer
On Error Resume Next
For oo = 20 To 0 Step -1
    comLx.RemoveItem oo
Next

If opt2.Value = True Then
    comLx.AddItem "外包"
    comLx.AddItem "水处理"
    comLx.Locked = False
End If
End Sub


Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0


If timZm = 1 Then    '业务添加

ElseIf timZm = 3 Then '签字
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
    If Dialog.Visible = True Then '更新事务列表
        Call mod1.refEnvent(1)
    End If
    If cmdQm(2).Caption <> "" And FMXC.Visible = True Then '业务员确认后，修改合同上的成本
        Call modNewHT.NewMQing
        Call modNewHT.NewB(Val(lblHtbh.Caption))
    End If
    Call QMBound(Val(lblBid.Caption))
ElseIf timZm = 8 Then '删除
    Me.Visible = False
    If FMXC.Visible = True Then
        FMXC.dtgFL.Col = 4

'''        FMXC.cmdW1.ToolTipText = ""
        FMXC.dtgFL.Row = 1: FMXC.dtgFL.Text = ""

    End If
    If Dialog.Visible = True Then
        Dialog.Enabled = True
        Dialog.ZOrder 0
    End If
End If
timQuit.Enabled = False
Me.Enabled = True
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
On Error Resume Next
timWait.Enabled = False
Me.Enabled = False
tt = "select cf,bz,bh,mm1,mm2,mt1,mt2 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then '业务添加
        Call MXBound(Val(lblBid.Caption))
        Select Case comLx.Text
        Case "主机维保"
            Call frmWBXT.Qing
            Call frmWBXT.Bound(mod1.WP.Fields("mm1").Value)
            frmWBXT.Show: frmWBXT.ZOrder 0
        Case "溴化锂"
            'Call frmWBXT1.Qing
            frmWBXT1.Show: frmWBXT1.ZOrder 0
        Case "小机末端空调箱保养"
            Call frmWBXT2.Qing
            Call frmWBXT2.Bound(mod1.WP.Fields("mm1").Value)
            frmWBXT2.Show: frmWBXT2.ZOrder 0
        End Select
        opt1.Value = False
        opt2.Value = False
        comLx.Locked = True
    ElseIf timZm = 2 Then '业务删除
        Call MXBound(Val(lblBid.Caption))
        txt2.Text = mod1.WP.Fields("mm2").Value
    ElseIf timZm = 3 Then '签字
        If OptT1.Value = True Then
            cmdQm(lblLc.Caption - 1).Caption = mod1.DName
            lblTm(lblLc.Caption - 1).Caption = mod1.DQda
        Else
            For ii = 0 To 2
                cmdQm(ii).Caption = ""
                lblTm(ii).Caption = ""
            Next
        End If
        lblLc.Caption = mod1.WP.Fields("mm1").Value
        lblFwid.Caption = mod1.WP.Fields("mm2").Value
        lblLcRen.Caption = mod1.WP.Fields("mt1").Value
        lblLcUid.Caption = mod1.WP.Fields("mt2").Value
        lblTX.Caption = "下一流程,将跳至" & lblQM(Val(lblLc.Caption) - 1).Caption & ": " & lblLcRen.Caption
    ElseIf timZm = 5 Then '表单保存
        cmdSave.Enabled = False
    ElseIf timZm = 6 Then '调整添加
        Call MXBound(Val(lblBid.Caption))
        frmTj.Visible = False
    End If

    timWait.Enabled = False

    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdJG.Enabled = False
    End If
    timWait.Enabled = False
    Me.Enabled = True
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdJG.Enabled = False
    End If
    Me.Enabled = True
    Exit Sub

End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub



Public Sub dtgPFF()
Dim oo As Integer
For oo = 1 To dtgP.Rows - 1
    dtgP.RowHeight(oo) = dtgP.RowHeight(0)
Next
dtgP.Clear
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "日期": dtgP.Col = 1: dtgP.Text = "姓名": dtgP.Col = 2: dtgP.Text = "职能": dtgP.Col = 3: dtgP.Text = "评审建议": dtgP.Col = 4: dtgP.Text = "审核":
dtgP.ColWidth(3) = 3000: dtgP.ColWidth(0) = 2000: dtgP.ColWidth(4) = 800
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub

Public Sub ji()
Dim hg As Single
Dim oo As Integer
Dim ii As Integer
On Error Resume Next
hg = 0
dtgN.Row = 1: dtgN.Col = 1
For oo = 1 To dtgN.Rows
    dtgN.Row = oo: dtgN.Col = 1
    hg = hg + Val(dtgN.Text)
    dtgN.Col = 2
    hg = hg + Val(dtgN.Text)
Next
txt2.Text = hg
End Sub

Public Sub QMBound(Bid As Long)
Dim Ra: Dim La
Dim ii As Integer: Dim oo As Integer
Dim tt As String
On Error Resume Next

tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & Bid & "' and yid=43 order by pid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2): dtgP.Rows = La + 20
dtgP.Clear
For oo = 1 To La + 1
    dtgP.Row = oo
    For ii = 0 To 5
        dtgP.Col = ii
        dtgP.Text = Ra(ii, oo - 1)
        If ii = 3 Then
            If Len(Ra(ii, oo - 1)) > 16 Then
                dtgP.RowHeight(oo) = UpInt(Len(Ra(ii, oo - 1)) / 16) * dtgP.RowHeight(oo)
            End If
        End If
        If ii = 4 Then
            If dtgP.Text = "True" Then
                dtgP.Text = "同意"
            ElseIf dtgP.Text = "False" Then
                dtgP.Text = "驳回"
            End If

        End If
    Next
Next
For oo = 1 To La + 1
    dtgP.Row = oo
    dtgP.Col = 4
            If dtgP.Text = "驳回" Then
                For ii = 0 To 5
                    dtgP.Col = ii
                    dtgP.CellForeColor = &HFF&
                Next
            End If
Next
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "日期": dtgP.Col = 1: dtgP.Text = "姓名": dtgP.Col = 2: dtgP.Text = "职能"
dtgP.Col = 3: dtgP.Text = "评审建议": dtgP.Col = 4: dtgP.Text = "通过否"



End Sub

Private Sub txtT3_Change()
txtT4.Text = Val(txtT2.Text) + Val(txtT3.Text) + Val(txtT6.Text)
End Sub

Private Sub txtT6_LostFocus()
txtT4.Text = Val(txtT2.Text) + Val(txtT3.Text) + Val(txtT6.Text)
End Sub


