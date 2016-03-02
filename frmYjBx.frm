VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmYjBx 
   BackColor       =   &H00C0FFC0&
   Caption         =   "奖金报销单"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11460
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   11460
   Begin VB.Frame frmLxr 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1035
      Left            =   3720
      TabIndex        =   45
      Top             =   960
      Width           =   7785
      Begin VB.TextBox txtPho5 
         BackColor       =   &H00BFFFE2&
         Height          =   270
         Left            =   2700
         TabIndex        =   63
         Top             =   690
         Width           =   1335
      End
      Begin VB.TextBox txtRen5 
         BackColor       =   &H00BFFFE2&
         Height          =   270
         Left            =   1020
         TabIndex        =   62
         Top             =   690
         Width           =   825
      End
      Begin VB.TextBox txtPho4 
         BackColor       =   &H00BFFFE2&
         Height          =   270
         Left            =   6270
         TabIndex        =   59
         Top             =   330
         Width           =   1395
      End
      Begin VB.TextBox txtRen4 
         BackColor       =   &H00BFFFE2&
         Height          =   270
         Left            =   4890
         TabIndex        =   58
         Top             =   330
         Width           =   825
      End
      Begin VB.TextBox txtPho3 
         BackColor       =   &H00BFFFE2&
         Height          =   270
         Left            =   2700
         TabIndex        =   55
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtRen3 
         BackColor       =   &H00BFFFE2&
         Height          =   270
         Left            =   1020
         TabIndex        =   54
         Top             =   360
         Width           =   825
      End
      Begin VB.TextBox txtRen2 
         BackColor       =   &H00BFFFE2&
         Height          =   270
         Left            =   4890
         TabIndex        =   51
         Top             =   30
         Width           =   825
      End
      Begin VB.TextBox txtPho2 
         BackColor       =   &H00BFFFE2&
         Height          =   270
         Left            =   6270
         TabIndex        =   50
         Top             =   30
         Width           =   1395
      End
      Begin VB.TextBox txtPho 
         BackColor       =   &H00BFFFE2&
         Height          =   270
         Left            =   2700
         TabIndex        =   49
         Top             =   60
         Width           =   1335
      End
      Begin VB.TextBox txtRen 
         BackColor       =   &H00BFFFE2&
         Height          =   270
         Left            =   1020
         TabIndex        =   48
         Text            =   "马晓聪大"
         Top             =   60
         Width           =   825
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "电话5"
         Height          =   195
         Left            =   2190
         TabIndex        =   65
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "受奖人5"
         Height          =   195
         Left            =   180
         TabIndex        =   64
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "电话4"
         Height          =   195
         Left            =   5760
         TabIndex        =   61
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "受奖人4"
         Height          =   195
         Left            =   4200
         TabIndex        =   60
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "电话3"
         Height          =   195
         Left            =   2160
         TabIndex        =   57
         Top             =   390
         Width           =   495
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "受奖人3"
         Height          =   195
         Left            =   180
         TabIndex        =   56
         Top             =   420
         Width           =   765
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "受奖人2"
         Height          =   195
         Left            =   4200
         TabIndex        =   53
         Top             =   90
         Width           =   645
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "电话2"
         Height          =   195
         Left            =   5760
         TabIndex        =   52
         Top             =   90
         Width           =   465
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "电话1"
         Height          =   195
         Left            =   2160
         TabIndex        =   47
         Top             =   90
         Width           =   495
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "受奖人1"
         Height          =   195
         Left            =   180
         TabIndex        =   46
         Top             =   120
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "审核"
      Height          =   765
      Left            =   10770
      Picture         =   "frmYjBx.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   4350
      Width           =   675
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   3660
      TabIndex        =   33
      Top             =   4080
      Visible         =   0   'False
      Width           =   6315
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   36
         Top             =   870
         Width           =   675
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   35
         Top             =   510
         Width           =   705
      End
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1305
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   300
         Width           =   4965
      End
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFC0C0&
      Caption         =   "保存"
      Height          =   795
      Left            =   10800
      Picture         =   "frmYjBx.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3540
      Width           =   645
   End
   Begin VB.CommandButton cmdWB 
      BackColor       =   &H00BFFFE2&
      Caption         =   "发付完毕"
      Height          =   315
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "确认此合同的奖金是否全部发放完毕"
      Top             =   1380
      Width           =   1005
   End
   Begin VB.TextBox txtBz 
      BackColor       =   &H00BFFFE2&
      Height          =   675
      Left            =   1530
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   990
      Width           =   2145
   End
   Begin VB.TextBox lblHtbh 
      BackColor       =   &H00BFFFE2&
      Height          =   270
      Left            =   8580
      TabIndex        =   26
      Top             =   90
      Width           =   2115
   End
   Begin VB.TextBox txtCXF 
      BackColor       =   &H00BFFFE2&
      Height          =   270
      Left            =   4740
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   720
      Width           =   1695
   End
   Begin VB.Frame frmHide 
      Caption         =   "frmHid"
      Height          =   1455
      Left            =   1080
      TabIndex        =   14
      Top             =   2250
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Timer timWait 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2550
         Top             =   270
      End
      Begin VB.Timer timQuit 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2670
         Top             =   900
      End
      Begin VB.Label lblUid 
         Caption         =   "lblUid"
         Height          =   255
         Left            =   3570
         TabIndex        =   23
         Top             =   900
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblYwy 
         Caption         =   "lblYwy"
         Height          =   285
         Left            =   3360
         TabIndex        =   22
         Top             =   510
         Width           =   765
      End
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   3780
         TabIndex        =   20
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Top             =   930
         Width           =   885
      End
      Begin VB.Label lblLcRen 
         Caption         =   "lblLcRen"
         Height          =   285
         Left            =   150
         TabIndex        =   18
         Top             =   420
         Width           =   795
      End
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   315
         Left            =   1050
         TabIndex        =   17
         Top             =   630
         Width           =   645
      End
      Begin VB.Label lblBm 
         Caption         =   "lblBm"
         Height          =   225
         Left            =   1020
         TabIndex        =   16
         Top             =   330
         Width           =   915
      End
      Begin VB.Label lblQy 
         Caption         =   "lblQy"
         Height          =   255
         Left            =   3150
         TabIndex        =   15
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00BFFFE2&
      Caption         =   "返回"
      Height          =   645
      Left            =   10770
      Picture         =   "frmYjBx.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5130
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   4065
      Left            =   0
      TabIndex        =   42
      Top             =   2070
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   7170
      _Version        =   393216
      BackColor       =   15728356
      ForeColor       =   0
      Rows            =   15
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   12648447
      ForeColorFixed  =   0
      BackColorSel    =   16744576
      BackColorBkg    =   15728356
      GridColorFixed  =   12640511
      GridColorUnpopulated=   12640511
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "审核栏"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   300
      TabIndex        =   44
      Top             =   1800
      Width           =   945
   End
   Begin VB.Label lblMx 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   7560
      TabIndex        =   41
      Top             =   2070
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.Label lblRen 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   8730
      TabIndex        =   40
      Top             =   1740
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblJR 
      BackStyle       =   0  'Transparent
      Caption         =   "受奖人"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7560
      TabIndex        =   39
      Top             =   1770
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblTX 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1470
      TabIndex        =   38
      Top             =   1800
      Width           =   5475
   End
   Begin VB.Label lblQFF 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1560
      TabIndex        =   31
      Top             =   750
      Width           =   1485
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "合同全款支付"
      Height          =   225
      Left            =   270
      TabIndex        =   30
      Top             =   750
      Width           =   1125
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "财务备注"
      Height          =   225
      Left            =   330
      TabIndex        =   28
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "手续费"
      Height          =   195
      Left            =   3900
      TabIndex        =   24
      Top             =   750
      Width           =   735
   End
   Begin VB.Label lblYid 
      Caption         =   "lblYid"
      Height          =   195
      Left            =   8760
      TabIndex        =   21
      Top             =   2130
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lblED 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      Height          =   225
      Left            =   8610
      TabIndex        =   13
      Top             =   420
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "收款额度"
      Height          =   195
      Left            =   7230
      TabIndex        =   12
      Top             =   420
      Width           =   1095
   End
   Begin VB.Label lblYf 
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   8610
      TabIndex        =   10
      Top             =   720
      Width           =   1635
   End
   Begin VB.Label lblCf 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      Height          =   225
      Left            =   4800
      TabIndex        =   9
      Top             =   420
      Width           =   1665
   End
   Begin VB.Label lblYj 
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      ForeColor       =   &H00C000C0&
      Height          =   225
      Left            =   1560
      TabIndex        =   8
      Top             =   420
      Width           =   1425
   End
   Begin VB.Label lblHtze 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      Height          =   195
      Left            =   4800
      TabIndex        =   7
      Top             =   90
      Width           =   1305
   End
   Begin VB.Label lblXmmc 
      BackStyle       =   0  'Transparent
      Caption         =   "Label6"
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   90
      Width           =   1515
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "支付金额"
      Height          =   225
      Left            =   7230
      TabIndex        =   5
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "曾付金额"
      Height          =   225
      Left            =   3690
      TabIndex        =   4
      Top             =   420
      Width           =   945
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Caption         =   "奖金总额"
      Height          =   225
      Left            =   270
      TabIndex        =   3
      Top             =   420
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "合同编号"
      Height          =   195
      Left            =   7230
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "合同金额"
      Height          =   195
      Left            =   3690
      TabIndex        =   1
      Top             =   90
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "项目名称"
      Height          =   195
      Left            =   270
      TabIndex        =   0
      Top             =   90
      Width           =   975
   End
End
Attribute VB_Name = "frmYjBx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer '数据提交后,由timWait执行的后续命令ID(1激活,2 签字
Private Sub cmdBack_Click()
Me.Visible = False
If frmBxBrow.Visible = True Then
    frmBxBrow.Enabled = True
    frmBxBrow.ZOrder 0
ElseIf Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf FMXC.Visible = True Then
    mod1.BTZ = 6
    FMXC.Enabled = True
    FMXC.ZOrder 0
End If
End Sub

Private Sub cmdDing_Click()
Dim tt As String
Dim CJ As Double
Dim CJB As Double
On Error Resume Next
If OptT1.Value = False And optT2.Value = False Then
    Exit Sub
End If
If OptT1.Value = True Then
    If Val(lblLc.Caption) = 3 Or Val(lblLc.Caption) = 1 Then '财务经理签字时,检查已收款
        If Val(Mid(Me.Caption, 34, 10)) / Val(lblHtze.Caption) < Val(lblED.Caption) / 100 Then
            ii = MsgBox("速达收款额度不足,是否同意签字?", vbQuestion + vbYesNo + vbDefaultButton2, "Hello!")
            If ii = vbNo Then Exit Sub
        End If
    End If
End If
If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
frmFX.Visible = False
timZm = 2 '签字
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "奖金"
    mod1.cmd.Parameters("@NBLX") = "新签字"
    mod1.cmd.Parameters("@bh") = lblYid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblYwy.Caption
    mod1.cmd.Parameters("@mt2") = lblUid.Caption
    mod1.cmd.Parameters("@mt3") = lblXmmc.Caption
    mod1.cmd.Parameters("@mt4") = lblHtbh.Text
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = ""
    mod1.cmd.Parameters("@mt11") = ""
    mod1.cmd.Parameters("@mt12") = ""
    mod1.cmd.Parameters("@mt13") = ""
    mod1.cmd.Parameters("@mt14") = ""
    mod1.cmd.Parameters("@mt15") = ""
    mod1.cmd.Parameters("@mt16") = ""
    mod1.cmd.Parameters("@mt17") = ""
    mod1.cmd.Parameters("@mt18") = ""
    mod1.cmd.Parameters("@mt19") = ""
    Select Case Val(lblLc.Caption)
        Case 1
            mod1.cmd.Parameters("@mt20") = "区域总监"
        Case 2
            mod1.cmd.Parameters("@mt20") = "销售总监"
        Case 3
            mod1.cmd.Parameters("@mt20") = "财务经理"""
        Case 4
            mod1.cmd.Parameters("@mt20") = "总经理"
        Case 5
            mod1.cmd.Parameters("@mt20") = "奖金签收"
        Case 6
            mod1.cmd.Parameters("@mt20") = "已支付确认"
    End Select
    '''''mod1.cmd.Parameters("@mt20") = lblQM(Val(lblLc.Caption) - 1).Caption
    Select Case Val(lblLc.Caption)
        Case 1
            mod1.cmd.Parameters("@mt21") = "开奖人"
        Case 2
            mod1.cmd.Parameters("@mt21") = "区域总监"
        Case 3
            mod1.cmd.Parameters("@mt21") = "销售总监"
        Case 4
            mod1.cmd.Parameters("@mt21") = "财务经理"
        Case 5
            mod1.cmd.Parameters("@mt21") = "总经理"
        Case 6
            mod1.cmd.Parameters("@mt21") = "奖金签收"
        Case 7
            mod1.cmd.Parameters("@mt21") = "已支付确认"
    End Select
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = txtQM.Text '评审建议
    mod1.cmd.Parameters("@mlt2") = txtBz.Text
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(lblLc.Caption)
    mod1.cmd.Parameters("@mm2") = Val(lblFwid.Caption)
    mod1.cmd.Parameters("@mm3") = Val(txtCXF.Text) '手续费
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
    mod1.cmd.Parameters("@mm16") = 0
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

Private Sub cmdNQ_Click()
''''If Val(lblLc.Caption) = 1 Then
''''    lblLcUid.Caption = mod1.DHid
''''End If
If lblLcUid.Caption <> mod1.DHid And mod1.Mname <> "马晓聪" Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If

If (txtRen.Text = "" Or txtPho.Text = "") And Val(lblLc.Caption) = 1 Then
    MsgBox "请输入受奖人和他的联系方式!"
    Exit Sub
End If


frmQm.Visible = True
OptT1.Value = False
optT2.Value = False
If lblLc.Caption = 1 Then
    optT2.Enabled = False
Else
    optT2.Enabled = True
End If
txtQM.Text = ""
End Sub



Private Sub cmdSave_Click()
Dim tt As String
If mod1.DName <> "乔继敏" And Val(lblLc.Caption) > 1 Then
    Exit Sub
End If
''''''tt = "update  yongjin set bz='" & txtBz.Text & "' where yid=" & Val(lblYid.Caption)
''''''Set mod1.HTP = CreateObject("adodb.recordset")
''''''mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText

Dim CJ As Double
Dim CJB As Double
On Error Resume Next

timZm = 3 '保存
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "奖金"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = lblYid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblYwy.Caption
    mod1.cmd.Parameters("@mt2") = lblUid.Caption
    mod1.cmd.Parameters("@mt3") = lblXmmc.Caption
    mod1.cmd.Parameters("@mt4") = lblHtbh.Text
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = ""
    mod1.cmd.Parameters("@mt11") = txtRen.Text
    mod1.cmd.Parameters("@mt12") = txtPho.Text
    mod1.cmd.Parameters("@mt13") = txtRen2.Text
    mod1.cmd.Parameters("@mt14") = txtPho2.Text
    mod1.cmd.Parameters("@mt15") = txtRen3.Text
    mod1.cmd.Parameters("@mt16") = txtPho3.Text
    mod1.cmd.Parameters("@mt17") = txtRen4.Text
    mod1.cmd.Parameters("@mt18") = txtPho4.Text
    mod1.cmd.Parameters("@mt19") = txtRen5.Text
    mod1.cmd.Parameters("@mt20") = txtPho5.Text
    mod1.cmd.Parameters("@mt21") = ""
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mlt2") = txtBz.Text
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = 0
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = Val(txtCXF.Text) '手续费
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
    mod1.cmd.Parameters("@mm16") = 0
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
        mod1.cmd.Parameters("@mb1") = 0 '拒绝
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


Private Sub cmdWb_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
ii = MsgBox("您是否确认此合同的奖金已经全部支付?", vbYesNo + vbInformation)
If ii = vbNo Then Exit Sub

If mod1.DName <> "乔继敏" Then
    Exit Sub
End If
''''''tt = "update  yongjin set bz='" & txtBz.Text & "' where yid=" & Val(lblYid.Caption)
''''''Set mod1.HTP = CreateObject("adodb.recordset")
''''''mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText


timZm = 5 '全款支付
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "奖金"
    mod1.cmd.Parameters("@NBLX") = "全款支付"
    mod1.cmd.Parameters("@bh") = lblYid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblHtbh.Text
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""
    mod1.cmd.Parameters("@mt6") = ""
    mod1.cmd.Parameters("@mt7") = ""
    mod1.cmd.Parameters("@mt8") = ""
    mod1.cmd.Parameters("@mt9") = ""
    mod1.cmd.Parameters("@mt10") = ""
    mod1.cmd.Parameters("@mt11") = ""
    mod1.cmd.Parameters("@mt12") = ""
    mod1.cmd.Parameters("@mt13") = ""
    mod1.cmd.Parameters("@mt14") = ""
    mod1.cmd.Parameters("@mt15") = ""
    mod1.cmd.Parameters("@mt16") = ""
    mod1.cmd.Parameters("@mt17") = ""
    mod1.cmd.Parameters("@mt18") = ""
    mod1.cmd.Parameters("@mt19") = ""
    mod1.cmd.Parameters("@mt20") = ""
    mod1.cmd.Parameters("@mt21") = ""
    mod1.cmd.Parameters("@mt22") = ""
    mod1.cmd.Parameters("@mt23") = ""
    mod1.cmd.Parameters("@mt24") = ""
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mlt2") = txtBz.Text
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = 0
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = Val(txtCXF.Text) '手续费
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
    mod1.cmd.Parameters("@mm16") = 0
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
        mod1.cmd.Parameters("@mb1") = 0 '拒绝
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


'''''Set mod1.cmd = createobject("adodb.command")
'''''mod1.cmd.ActiveConnection = mod1.CC
'''''mod1.cmd.CommandText = "newYjFF"
'''''mod1.cmd.CommandType = adCmdStoredProc
'''''mod1.cmd.Parameters("@htbh").Value = lblHtbh.Text
'''''mod1.cmd.Execute
'''''tt = mod1.cmd.Parameters("@errinf").Value
'''''Set cmd = Nothing
'''''
'''''If tt = "执行成功!" Then
'''''    MsgBox ("OK,合同评审单的实际奖金已经完全归入!")
'''''    cmdWb.Visible = False
'''''    frmYjBx.lblQFF.Caption = "全款支付完毕"
'''''    frmYjBx.lblQFF.ForeColor = &HFF&
'''''Else
'''''    MsgBox ("网络错误,请退出程序再试一次,或与马晓聪联系!")
'''''End If
End Sub

Private Sub Form_Click()
frmQm.Visible = False
lblMx.Visible = False
End Sub


Private Sub Form_DblClick()
''''''Dim ii As Integer
''''''Dim oo As Integer
''''''Dim Qlabel As String
''''''On Error Resume Next
''''''frmQm.Visible = False
''''''If lblQM(0).Visible = False And mod1.BmJl = True Then
''''''    tt = "select sum(zfu) as zfu ,sum(kf) as kf from yjz where htbh='" & lblHtbh.Text & "'"
''''''    Set mod1.HTP = CreateObject("adodb.recordset")
''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''    If mod1.HTP.Fields("zfu").Value > 0 And mod1.HTP.Fields("kf").Value > 0 Then
'''''''        MsgBox ("旧版豪曼信息中显示,曾经支付过金额" & mod1.HTP.Fields("zfu").Value & ",请与周工联系!")
'''''''        Exit Sub
'''''''    End If
''''''    ii = MsgBox("是否要激活此单的运作?", vbQuestion + vbYesNo + vbDefaultButton2, "您好!")
''''''    If ii = vbYes Then
''''''
'''''''''''''''        For oo = 0 To 6
'''''''''''''''            Select Case oo
'''''''''''''''            Case 0
'''''''''''''''                Qlabel = "开奖"
'''''''''''''''            Case 1
'''''''''''''''                Qlabel = "销售总监"
'''''''''''''''            Case 2
'''''''''''''''                Qlabel = "财务确认"
'''''''''''''''            Case 3
'''''''''''''''                Qlabel = "总经理"
'''''''''''''''            Case 4
'''''''''''''''                Qlabel = "财务总监"
'''''''''''''''            Case 5
'''''''''''''''                Qlabel = "签收"
'''''''''''''''            Case 6
'''''''''''''''                Qlabel = "已支付"
'''''''''''''''            End Select
'''''''''''''''            lblQM(oo).Caption = Qlabel
'''''''''''''''            frmYjBx.lblQM(oo).Visible = True
'''''''''''''''            frmYjBx.lblTm(oo).Visible = True
'''''''''''''''            frmYjBx.cmdQm(oo).Visible = True
'''''''''''''''            tt = "insert into qmrz (qlabel,btz,qdbh,zid) values ('" & Qlabel & "',23,'" & lblYid.Caption & "'," & (oo + 1) & ")"
'''''''''''''''            Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''''''''            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'''''''''''''''
'''''''''''''''        Next
'''''''''''''''        lblYwy.Caption = mod1.DName
'''''''''''''''        lblUid.Caption = mod1.DHid
'''''''''''''''        lblLcRen.Caption = mod1.DName
'''''''''''''''        lblLcUid.Caption = mod1.DHid
'''''''''''''''        Call mod1.EnventAdd("奖金", lblXmmc.Caption, lblLcRen.Caption, lblLcUid.Caption, lblYid.Caption, lblQM(0).Caption, "", "", mod1.DName, mod1.DHid, 0, lblYid.Caption)
'''''''''''''''
'''''''''''''''        tt = "update yongjin set ywy='" & mod1.DName & "',uid='" & mod1.DHid & "',lc=1,lcren='" & mod1.DName & "',lcuid='" & _
'''''''''''''''        mod1.DHid & "',fwid=" & Val(lblFwid.Caption) & " where yid=" & Val(lblYid.Caption)
'''''''''''''''        Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''''''''        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'''''''''''''''        lblLc.Caption = 1
'''''''''''''''        frmBxBrow.adoYj.Requery
'''''''''''''''        Set frmBxBrow.dtgYj.DataSource = frmBxBrow.adoYj
''''''timZm = 1 '激活
''''''    Set mod1.cmd = createobject("adodb.command")
''''''    mod1.cmd.ActiveConnection = mod1.cc
''''''    mod1.cmd.CommandText = "MLAdd"
''''''    mod1.cmd.CommandType = adCmdStoredProc
''''''    mod1.cmd.Parameters("@zid") = 0
''''''    mod1.cmd.Parameters("@errch") = ""
''''''    mod1.cmd.Parameters("@NB") = "奖金"
''''''    mod1.cmd.Parameters("@NBLX") = "激活"
''''''    mod1.cmd.Parameters("@bh") = lblYid.Caption
''''''    mod1.cmd.Parameters("@ywy") = mod1.DName
''''''    mod1.cmd.Parameters("@uid") = mod1.DHid
''''''    mod1.cmd.Parameters("@mt1") = lblXmmc.Caption
''''''    mod1.cmd.Parameters("@mt2") = ""
''''''    mod1.cmd.Parameters("@mt3") = ""
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
''''''    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
''''''    mod1.cmd.Execute
''''''    mod1.Zid = mod1.cmd.Parameters("@zid").Value
''''''    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
''''''        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
''''''        cmdDing.Enabled = False
''''''        Exit Sub
''''''    Else '提交成功,等待系统中心处理数据
''''''        Me.Enabled = False
''''''        frmWaitA.Visible = True
''''''        frmWaitA.Timer2.Enabled = False
''''''
''''''        frmWaitA.ZOrder 0
''''''        frmWaitA.Timer2.Enabled = True
''''''        timWait.Enabled = True
''''''    End If
''''''
''''''
''''''Set mod1.cmd = Nothing
''''''    End If
''''''End If
End Sub

Private Sub Form_Load()
Dim oo As Integer
Me.Width = 11580
Me.Height = 6390

frmQm.Left = 3750
frmQm.Top = 4080
dtgP.Left = 0
dtgP.Top = 2070
End Sub

Public Sub yjBXQing()
lblXmmc.Caption = ""
lblYj.Caption = ""
lblHtze.Caption = ""
lblCf.Caption = ""
lblHtbh.Text = ""
lblED.Caption = ""
lblYf.Caption = ""
lblBM.Caption = ""
lblQy.Caption = ""
lblLcRen.Caption = ""
lblLcUid.Caption = ""
lblLc.Caption = ""
lblYwy.Caption = ""
lblUid.Caption = ""
lblFwid.Caption = ""
txtCXF.Text = ""
lblQFF.Caption = ""
lblQFF.ForeColor = &H80000012
txtBz.Text = ""
frmQm.Visible = False
txtQM.Text = ""
OptT1.Value = False
optT2.Value = False
lblTX.Visible = False
lblTX.Caption = ""
lblRen.Caption = ""
lblRen.ToolTipText = ""
lblRen.Tag = 0
lblMx.Caption = ""
lblMx.Visible = False
frmLxr.Visible = False
txtRen.Text = ""
txtPho.Text = ""
txtRen2.Text = ""
txtPho2.Text = ""
txtRen3.Text = ""
txtPho3.Text = ""
txtRen4.Text = ""
txtPho4.Text = ""
txtRen5.Text = ""
txtPho5.Text = ""
cmdSave.Visible = False
End Sub



Private Sub lblMx_Click()
lblMx.Visible = False
End Sub

Private Sub lblRen_Click()
lblMx.Visible = True
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub timQuit_Timer()
Dim oo As Integer
Dim ii As Integer
On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
Dim tt As String
If timZm = 1 Then '激活
        Call QMBound(Val(lblYid.Caption))
        lblYwy.Caption = mod1.DName
        lblUid.Caption = mod1.DHid
        lblLcRen.Caption = mod1.DName
        lblLcUid.Caption = mod1.DHid
        lblLc.Caption = 1
        frmBxBrow.adoYj.Requery
        Set frmBxBrow.dtgYJ.DataSource = frmBxBrow.adoYj
ElseIf timZm = 2 Then '签字
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
    If Dialog.Visible = True Then
        Call mod1.refEnvent(1)
    End If
ElseIf timZm = 3 Then '保存
    If frmBxBrow.Visible = True Then
        frmBxBrow.dtgYJ.Col = 10
        frmBxBrow.dtgYJ.Text = txtBz.Text
    End If
    cmdSave.Visible = False
ElseIf timZm = 5 Then
    MsgBox ("OK,合同评审单的实际奖金已经完全归入!")
    cmdWb.Visible = False
    frmYjBx.lblQFF.Caption = "全款支付完毕"
    frmYjBx.lblQFF.ForeColor = &HFF&
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
    If timZm = 1 Then '激活
        Label1.Caption = Label1.Caption
        lblFwid.Caption = mod1.WP.Fields("mm1").Value

    ElseIf timZm = 2 Then '签字
'''''        If OptT1.Value = True Then
'''''            cmdQm(lblLc.Caption - 1).Caption = mod1.DName
'''''            lblTm(lblLc.Caption - 1).Caption = mod1.DQda
'''''        Else
'''''            For oo = 0 To 6
'''''                cmdQm(oo).Caption = ""
'''''                lblTm(oo).Caption = ""
'''''            Next
'''''        End If
        lblLc.Caption = mod1.WP.Fields("mm1").Value
        lblFwid.Caption = mod1.WP.Fields("mm2").Value
        lblLcRen.Caption = mod1.WP.Fields("mt1").Value
        lblLcUid.Caption = mod1.WP.Fields("mt2").Value
        lblTX.Caption = "下一流程,将跳至:" & lblLcRen.Caption
        Call QMBound(lblYid.Caption)
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    timWait.Enabled = False
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 2 Then
        cmdSave.Enabled = False
    ElseIf timZm = 11 Then
        txtHtbh.Text = ""
        lblHtxz.Caption = ""
    End If
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 2 Then
        cmdSave.Enabled = False
    ElseIf timZm = 11 Then
        txtHtbh.Text = ""
        lblHtxz.Caption = ""
    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub



Public Sub QMBound(Yid As Long)
Dim Ra: Dim La
Dim ii As Integer: Dim oo As Integer
Dim tt As String
On Error GoTo YJBX2

tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & Yid & "' and yid=68 order by pid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.BOF = False Then
Ra = mod1.HTP.GetRows
End If
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next
La = UBound(Ra, 2): dtgP.Rows = La + 20
'If La = 0 Then Exit Sub
dtgP.Visible = False
Call dtgPFF
For oo = 1 To La + 1
    dtgP.Row = oo
    For ii = 0 To 5
        dtgP.Col = ii
        dtgP.Text = Ra(ii, oo - 1)
            DH = 255 * mod1.HH(dtgP.Text, UpInt(dtgP.CellWidth / 200))
            If DH > dtgP.RowHeight(dtgP.Row) Then
                dtgP.RowHeight(dtgP.Row) = DH
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

dtgP.TopRow = 1
dtgP.Visible = True
Exit Sub
YJBX2:
MsgBox "出错!"
End
End Sub

Public Sub dtgPFF()
Dim oo As Integer
For oo = 1 To dtgP.Rows - 1
    dtgP.RowHeight(oo) = dtgP.RowHeight(0)
Next
dtgP.Clear
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "日期": dtgP.Col = 1: dtgP.Text = "姓名": dtgP.Col = 2: dtgP.Text = "职能": dtgP.Col = 3: dtgP.Text = "评审建议": dtgP.Col = 4: dtgP.Text = "审核":
dtgP.ColWidth(0) = 2220
dtgP.ColWidth(1) = 1800
dtgP.ColWidth(2) = 1000
 dtgP.ColWidth(3) = 3530: dtgP.ColWidth(4) = 975
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub
Public Sub Lren(Hid As Long)
Dim tt As String
Dim Ra
Dim Sex As String
Dim khZw As String
Dim khDpho As String
Dim khMob As String
On Error Resume Next
tt = "SELECT dbo.khRen.khMan, dbo.khRen.rId, dbo.khRen.khSex, dbo.khRen.khZw, dbo.khRen.khDpho, dbo.khRen.khMob, dbo.htPing.Hid FROM dbo.htPing INNER JOIN dbo.khRen ON dbo.htPing.rid = dbo.khRen.rId" & _
    " where dbo.htping.hid=" & Hid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
lblRen.Caption = Ra(0, 0)
lblRen.Tag = Ra(1, 0)
Sex = Ra(2, 0)
khZw = Ra(3, 0)
khDpho = Ra(4, 0)
khMob = Ra(5, 0)
'lblMx.Caption = "性别：     " & Sex & Chr(13) & Chr(10) & "职务：     " & khZw & Chr(13) & Chr(10) & "电话：     " & khDpho & Chr(13) & Chr(10) & "手机：     " & khMob
lblMx.Caption = "电话：     " & khDpho & Chr(13) & Chr(10) & "手机：     " & khMob
End Sub

Public Sub LrenH(Ra)
Dim tt As String

Dim Sex As String
Dim khZw As String
Dim khDpho As String
Dim khMob As String
On Error Resume Next
'''''tt = "SELECT dbo.khRen.khMan, dbo.khRen.rId, dbo.khRen.khSex, dbo.khRen.khZw, dbo.khRen.khDpho, dbo.khRen.khMob, dbo.htPing.Hid FROM dbo.htPing INNER JOIN dbo.khRen ON dbo.htPing.rid = dbo.khRen.rId" & _
'''''    " where dbo.htping.htbh='" & Htbh & "'"
'''''Set mod1.HTP = CreateObject("adodb.recordset")
'''''mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''Ra = mod1.HTP.GetRows
'''''mod1.HTP.Close
'''''Set mod1.HTP = Nothing
lblRen.Caption = Ra(0, 0)
lblRen.Tag = Ra(1, 0)
Sex = Ra(2, 0)
khZw = Ra(3, 0)
khDpho = Ra(4, 0)
khMob = Ra(5, 0)
'lblMx.Caption = "性别：     " & Sex & Chr(13) & Chr(10) & "职务：     " & khZw & Chr(13) & Chr(10) & "电话：     " & khDpho & Chr(13) & Chr(10) & "手机：     " & khMob
lblMx.Caption = "电话：     " & khDpho & Chr(13) & Chr(10) & "手机：     " & khMob
End Sub

Public Sub Bound(Yid As Long)
Dim tt As String
Dim oo As Integer
Dim Pwf As Boolean
Dim QFF As Boolean '合同全款支付否
Dim Ny As Single '已支付的奖金总额(新版中的,不是梅花档案中的)
Dim Xmmc As String
Dim Ra, Rb, RC, RD, RE

Call frmYjBx.yjBXQing
On Error GoTo YJERRB
tt = "declare @htbh nvarchar(50);" & _
    "select * from newyjht where yid=" & Yid & ";" & _
    "select @htbh=合同编号 from newyjht where yid=" & Yid & ";" & _
    "select yj from htping where htbh=@htbh;" & _
    "SELECT dbo.khRen.khMan, dbo.khRen.rId, dbo.khRen.khSex, dbo.khRen.khZw, dbo.khRen.khDpho, dbo.khRen.khMob," & _
    "dbo.htPing.Hid FROM dbo.htPing INNER JOIN dbo.khRen ON dbo.htPing.rid = dbo.khRen.rId" & _
    " where dbo.htping.htbh=@htbh;" & _
    "select sum(应付)+sum(cxf) from newyjht where 合同编号=@htbh and 支付否=1;" & _
    "Select sum(zFu) as zfu from yjz where htbh=@htbh;" & _
     "select sum(amount) from SDV_ChargeA where htbh=@htbh"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText

If mod1.HTP.BOF = True Then
    MsgBox "出错!可能是此奖金单对应的合同评审单已经被删除,请仔细确认!"
    End
Else
    Ra = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = True Then
    MsgBox "出错!"
    End
Else
    Rb = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = True Then
'    MsgBox "出错!"
'    End
Else
    RC = mod1.HTP.GetRows
    
End If
Set mod1.HTP = mod1.HTP.NextRecordset
RD = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RE = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    Rf = mod1.HTP.GetRows
End If
mod1.HTP.Close

Set mod1.HTP = Nothing
QFF = False
On Error Resume Next
frmYjBx.lblQy.Caption = Ra(0, 0)
frmYjBx.lblBM.Caption = Ra(1, 0)
frmYjBx.lblXmmc.Caption = Ra(2, 0)
frmYjBx.lblHtbh.Text = Ra(4, 0)
frmYjBx.lblHtze.Caption = Ra(3, 0)
frmYjBx.lblYf.Caption = Ra(6, 0)
frmYjBx.lblED.Caption = Ra(5, 0)
frmYjBx.lblYwy.Caption = Ra(13, 0)
frmYjBx.lblUid.Caption = Ra(14, 0)
QFF = Ra(22, 0)
If QFF = True Then
    frmYjBx.lblQFF.Caption = "全款支付完毕"
    frmYjBx.lblQFF.ForeColor = &HFF&
Else
    frmYjBx.lblQFF.Caption = "未完成"
End If

frmYjBx.txtRen.Text = Ra(24, 0)
frmYjBx.txtPho.Text = Ra(25, 0)
frmYjBx.txtRen2.Text = Ra(26, 0)
frmYjBx.txtPho2.Text = Ra(27, 0)
frmYjBx.txtRen3.Text = Ra(28, 0)
frmYjBx.txtPho3.Text = Ra(29, 0)
frmYjBx.txtRen4.Text = Ra(30, 0)
frmYjBx.txtPho4.Text = Ra(31, 0)
frmYjBx.txtRen5.Text = Ra(32, 0)
frmYjBx.txtPho5.Text = Ra(33, 0)
frmYjBx.lblYid.Caption = Ra(10, 0)

frmYjBx.lblYwy.Caption = Ra(13, 0)
frmYjBx.lblUid.Caption = Ra(14, 0)
frmYjBx.lblLc.Caption = Ra(15, 0)
frmYjBx.lblLcRen.Caption = Ra(16, 0)
frmYjBx.lblLcUid.Caption = Ra(17, 0)
frmYjBx.lblFwid.Caption = Ra(18, 0)
frmYjBx.txtCXF.Text = Ra(20, 0)
Pwf = Ra(21, 0)

frmYjBx.txtBz.Text = Ra(9, 0)
Ny = 0

frmYjBx.lblYj.Caption = Rb(0, 0)

Call frmYjBx.LrenH(RC)

'''''tt = "select sum(应付)+sum(cxf) from newyjht where 合同编号='" & frmYjBx.lblHtbh.Text & "' and 支付否=1"
'''''Set mod1.HTP = CreateObject("adodb.recordset")
'''''mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText



'检查梅花档案中的曾经支付
'实际表
'''''tt = "Select sum(zFu) as zfu from yjz where htbh='" & frmYjBx.lblHtbh.Text & "'"
'''''mod1.HTT.Close
'''''mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''If IsNull(mod1.HTP.Fields(0).Value) = True Then
''''''    frmYjBx.lblCf.Caption = 0
''''''Else

If IsNull(RD(0, 0)) = True Then
    Ny = 0
Else
    Ny = RD(0, 0)
End If
If IsNull(RE(0, 0)) = True Then
    frmYjBx.lblCf.Caption = Ny
Else
    frmYjBx.lblCf.Caption = Ny + RE(0, 0)
End If

'End If

'''''''For oo = 0 To 6
'''''''    frmYjBx.lblTm(oo).Caption = ""
'''''''    frmYjBx.cmdQm(oo).Caption = ""
'''''''    frmYjBx.lblQM(oo).Visible = False
'''''''    frmYjBx.lblTm(oo).Visible = False
'''''''    frmYjBx.cmdQm(oo).Visible = False
'''''''Next
'''''''
''''''''判断有无签字按钮,若没有,则添加
'''''''If frmYjBx.lblYwy.Caption <> "" Then
'''''''    tt = "select * from qmrz where btz=23 and qdbh='" & frmYjBx.lblYid.Caption & "' order by zid"
'''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''
'''''''    mod1.HTP.MoveFirst
'''''''    For oo = 0 To 6
'''''''        frmYjBx.lblQM(oo).Caption = mod1.HTP.Fields("qLabel").Value
'''''''        If mod1.HTP.Fields("xf").Value = True Then
'''''''            frmYjBx.cmdQm(oo).Caption = mod1.HTP.Fields("qren").Value
'''''''            If frmYjBx.cmdQm(oo).Caption = "南京办经理" Then
'''''''                frmYjBx.cm   dQm(oo).Caption = "南京办经理"
'''''''            End If
'''''''            frmYjBx.lblTm(oo).Caption = mod1.HTP.Fields("qrq").Value
'''''''        End If
'''''''        frmYjBx.cmdQm(oo).Visible = True
'''''''        frmYjBx.lblQM(oo).Visible = True
'''''''        frmYjBx.lblTm(oo).Visible = True
'''''''        mod1.HTP.MoveNext
'''''''    Next
'''''''    If frmYjBx.lblQM(5).Caption = "已支付" Then
'''''''        frmYjBx.lblQM(6).Visible = False
'''''''        frmYjBx.cmdQm(6).Visible = False
'''''''        frmYjBx.lblTm(6).Visible = False
'''''''    End If
'''''''    If Pwf = True And frmYjBx.cmdQm(5).Caption = "" And frmYjBx.cmdQm(6).Visible = False Then '已支付显示
'''''''        frmYjBx.cmdQm(5).Caption = frmYjBx.cmdQm(2).Caption
'''''''        frmYjBx.lblTm(5).Caption = frmYjBx.lblTm(4).Caption
'''''''    End If
'''''''
'''''''Else
'''''''
'''''''End If
'速达收款
 If Rf(0, 0) > 0 Then
    Me.Caption = "奖金报销单                  速达已收款:    " & Rf(0, 0)
 Else
    Me.Caption = "奖金报销单"
 End If

Call Me.dtgPFF
Call Me.QMBound(Yid)
If QFF = False And mod1.DName = "乔继敏" And Pwf = True Then
    frmYjBx.cmdWb.Visible = True
Else
    frmYjBx.cmdWb.Visible = False
End If
'旧的合同没有执行审核人,系统将搜索销售经理来指定
If (Val(lblLc.Caption) = 0 Or Val(lblLc.Caption) = 1) And lblLcRen.Caption = "" Then
    tt = "select xywy,xuid from htping where htbh='" & lblHtbh.Text & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    Set mod1.HTP = Nothing
    tt = "select ggl from worker where userid='" & Ra(1, 0) & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rb = mod1.HTP.GetRows
    Set mod1.HTP = Nothing
    If Rb(0, 0) = mod1.DHid Then
        lblLcRen.Caption = mod1.DName
        lblLcUid.Caption = mod1.DHid
    Else
        lblLcRen.Caption = Ra(0, 0)
        lblLcUid.Caption = Ra(1, 0)
    End If
End If

If Pwf = False Then
    lblTX.Visible = True
    lblTX.Caption = "流程至: " & lblLcRen.Caption
Else
    lblTX.Visible = False
End If

frmBxBrow.Enabled = False
frmYjBx.Show
frmYjBx.ZOrder 0
frmYjBx.OptT1.Value = False
frmYjBx.optT2.Value = False
''''If Val(lblLc.Caption) = 1 Then
''''    frmLxr.Visible = True
''''Else
''''    frmLxr.Visible = False
''''End If
frmLxr.Visible = False
If (Val(lblLc.Caption) = 1 And lblLcUid.Caption = mod1.DHid) Or mod1.DName = "乔继敏" Or mod1.DName = "宋晓炯" Or mod1.DName = "沈维" Or mod1.Mname = "马晓聪" Then
    cmdSave.Visible = True
    frmLxr.Visible = True
End If


Exit Sub
YJERRB:
MsgBox "出错!"
End
End Sub
