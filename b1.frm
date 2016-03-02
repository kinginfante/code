VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form b1 
   Caption         =   "上海豪曼制冷空调服务有限公司"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdBB 
      Caption         =   "报表"
      Height          =   615
      Left            =   11580
      Picture         =   "b1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   8580
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1365
      Left            =   9180
      TabIndex        =   47
      Top             =   150
      Visible         =   0   'False
      Width           =   2505
      Begin VB.Label lblFwid 
         Caption         =   "lblFwid"
         Height          =   255
         Left            =   1230
         TabIndex        =   50
         Top             =   0
         Width           =   885
      End
      Begin VB.Label lblLcUid 
         Caption         =   "lblLcUid"
         Height          =   285
         Left            =   30
         TabIndex        =   49
         Top             =   810
         Width           =   885
      End
      Begin VB.Label lblLc 
         Caption         =   "lblLc"
         Height          =   315
         Left            =   0
         TabIndex        =   48
         Top             =   390
         Width           =   645
      End
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   3270
      TabIndex        =   42
      Top             =   5520
      Visible         =   0   'False
      Width           =   6315
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   45
         Top             =   870
         Width           =   675
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   44
         Top             =   480
         Width           =   705
      End
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1365
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   300
         Width           =   4965
      End
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H008080FF&
      Caption         =   "新建"
      Height          =   345
      Left            =   14460
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   810
      Width           =   555
   End
   Begin MSComCtl2.DTPicker txtM 
      Height          =   345
      Left            =   12390
      TabIndex        =   38
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy年MM月"
      Format          =   56492035
      CurrentDate     =   39415
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11970
      Top             =   7740
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12750
      Top             =   7740
   End
   Begin VB.Frame frmEd 
      Caption         =   "编辑栏"
      Height          =   1605
      Left            =   3660
      TabIndex        =   26
      Top             =   7590
      Visible         =   0   'False
      Width           =   7785
      Begin VB.CommandButton cmdDao 
         Caption         =   "导入上月任务"
         Height          =   345
         Left            =   4290
         TabIndex        =   53
         Top             =   1230
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtp5 
         Height          =   285
         Left            =   1890
         TabIndex        =   41
         Top             =   930
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56492033
         CurrentDate     =   39415
      End
      Begin VB.CommandButton cmdGx 
         Caption         =   "更新"
         Height          =   375
         Left            =   6060
         TabIndex        =   36
         Top             =   1200
         Width           =   525
      End
      Begin VB.CommandButton cmdQing 
         Caption         =   "清空"
         Height          =   375
         Left            =   6570
         TabIndex        =   35
         Top             =   1200
         Width           =   525
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加"
         Height          =   375
         Left            =   5520
         TabIndex        =   34
         Top             =   1200
         Width           =   525
      End
      Begin VB.CommandButton cmdDe 
         Caption         =   "删除"
         Height          =   375
         Left            =   7080
         TabIndex        =   33
         Top             =   1200
         Width           =   525
      End
      Begin VB.TextBox txt3 
         Height          =   270
         Left            =   5790
         TabIndex        =   32
         Top             =   870
         Width           =   1725
      End
      Begin VB.TextBox txt2 
         Height          =   270
         Left            =   1890
         TabIndex        =   31
         Tag             =   "50"
         ToolTipText     =   "50"
         Top             =   540
         Width           =   5625
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Left            =   1890
         TabIndex        =   30
         Tag             =   "50"
         ToolTipText     =   "50"
         Top             =   210
         Width           =   5625
      End
      Begin VB.Label Label9 
         Caption         =   "完成时间："
         Height          =   225
         Left            =   900
         TabIndex        =   40
         Top             =   990
         Width           =   915
      End
      Begin VB.Label lblGid 
         Caption         =   "lblGid"
         Height          =   195
         Left            =   540
         TabIndex        =   37
         Top             =   1290
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "重要性基数："
         Height          =   195
         Left            =   4620
         TabIndex        =   29
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "工作目标："
         Height          =   195
         Left            =   900
         TabIndex        =   28
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "月度工作计划内容："
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdR 
      Caption         =   "->"
      Height          =   315
      Left            =   14010
      TabIndex        =   24
      Top             =   870
      Width           =   405
   End
   Begin VB.CommandButton cmdL 
      Caption         =   "<-"
      Height          =   315
      Left            =   12000
      TabIndex        =   23
      Top             =   870
      Width           =   405
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "查询"
      Height          =   585
      Left            =   12330
      Picture         =   "b1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8580
      Width           =   825
   End
   Begin VB.CommandButton cmdZuan 
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14790
      TabIndex        =   21
      Top             =   0
      Width           =   435
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "提交"
      Enabled         =   0   'False
      Height          =   585
      Left            =   13860
      Picture         =   "b1.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8580
      Width           =   675
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14580
      Picture         =   "b1.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8580
      Width           =   585
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   585
      Left            =   13170
      Picture         =   "b1.frx":0FF0
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8580
      Width           =   645
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgB1 
      Height          =   4515
      Left            =   660
      TabIndex        =   17
      Top             =   2910
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   7964
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtGzgy 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Tag             =   "200"
      Text            =   "b1.frx":12FA
      Top             =   1530
      Width           =   12945
   End
   Begin VB.TextBox txtZw 
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
      Left            =   9150
      TabIndex        =   15
      Text            =   "Text3"
      Top             =   870
      Width           =   1845
   End
   Begin VB.TextBox txtBm 
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
      Left            =   5310
      TabIndex        =   14
      Text            =   "Text2"
      Top             =   870
      Width           =   2295
   End
   Begin VB.TextBox txtYwy 
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
      Left            =   2010
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   870
      Width           =   1485
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      Top             =   8280
      Width           =   1485
   End
   Begin VB.CommandButton cmdMQm 
      Height          =   345
      Index           =   0
      Left            =   510
      TabIndex        =   7
      Top             =   8280
      Width           =   1425
   End
   Begin VB.CommandButton cmdPje 
      Caption         =   "评审建议"
      Height          =   1095
      Left            =   90
      TabIndex        =   6
      Top             =   7980
      Width           =   345
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
      Left            =   600
      TabIndex        =   51
      Top             =   7560
      Width           =   5475
   End
   Begin VB.Label lblKid 
      Caption         =   "lblKid"
      Height          =   225
      Left            =   900
      TabIndex        =   25
      Top             =   330
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblMQM 
      Caption         =   "被考核员工签名"
      Height          =   225
      Index           =   1
      Left            =   2100
      TabIndex        =   12
      Top             =   8010
      Width           =   1365
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   2040
      TabIndex        =   11
      Top             =   8700
      Width           =   1485
   End
   Begin VB.Label lblMTm 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   510
      TabIndex        =   10
      Top             =   8700
      Width           =   1425
   End
   Begin VB.Label lblMQM 
      Caption         =   "部门负责人签名"
      Height          =   225
      Index           =   0
      Left            =   570
      TabIndex        =   9
      Top             =   8010
      Width           =   1275
   End
   Begin VB.Label Label6 
      Caption         =   "本月工作概要"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   660
      TabIndex        =   5
      Top             =   1560
      Width           =   1125
   End
   Begin VB.Label Label5 
      Caption         =   "月份"
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
      Left            =   11460
      TabIndex        =   4
      Top             =   900
      Width           =   585
   End
   Begin VB.Label Label4 
      Caption         =   "岗位"
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
      Left            =   8100
      TabIndex        =   3
      Top             =   900
      Width           =   885
   End
   Begin VB.Label Label3 
      Caption         =   "部门"
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
      Left            =   4140
      TabIndex        =   2
      Top             =   900
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "姓名："
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
      Left            =   720
      TabIndex        =   1
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "员工月度工作计划表"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4890
      TabIndex        =   0
      Top             =   150
      Width           =   3405
   End
End
Attribute VB_Name = "b1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public M1 As ADODB.Recordset
Dim BF As Boolean '打开表单成功否
Dim timZm As Integer '1员工表1新建,2计划编辑3保存5签字6导入上月
Dim yGGl As String '上级管理者
Dim adoMM As ADODB.Recordset
Public Sub OAn()
Dim adoMM As ADODB.Recordset
Dim zt As String
Dim oo As Integer
On Error Resume Next
Set adoMM = New ADODB.Recordset


      zt = "qmrzOpen(" & 63 & ",'" & lblKid.Caption & "')"
      adoMM.Close
      adoMM.Open zt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc



      If IsNull(adoMM.RecordCount) = True Then
        MsgBox ("出错7")
        Exit Sub
      End If
      'MsgBox ("jj")
      If adoMM.RecordCount > 0 Then
         adoMM.MoveFirst
         cmdMQm(0).Visible = True
         lblMQM(0).Visible = True
         lblMTm(0).Visible = True
                  lblMQM(0).Caption = adoMM.Fields("QLabel").Value
            If adoMM.Fields("xf").Value = True Then
                cmdMQm(0).Caption = adoMM.Fields("Qren").Value
                lblMTm(0).Caption = adoMM.Fields("QRQ").Value
             End If
         cmdMQm(0).Tag = adoMM.Fields("zid").Value
         adoMM.MoveNext
         For oo = 1 To adoMM.RecordCount - 1
           lblMQM(oo).Caption = adoMM.Fields("QLabel").Value
            If adoMM.Fields("xf").Value = True Then
                cmdMQm(oo).Caption = adoMM.Fields("Qren").Value
                lblMTm(oo).Caption = adoMM.Fields("QRQ").Value
           End If
           cmdMQm(oo).Tag = adoMM.Fields("zid").Value
           adoMM.MoveNext
        Next

     End If
     
     '表2
     
Set adoMM = New ADODB.Recordset


      zt = "qmrzOpen(" & 66 & ",'" & lblKid.Caption & "')"
      adoMM.Close
      adoMM.Open zt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc



      If IsNull(adoMM.RecordCount) = True Then
        MsgBox ("出错7")
        Exit Sub
      End If
      'MsgBox ("jj")
      If adoMM.RecordCount > 0 Then
         adoMM.MoveFirst
         b2.cmdMQm(0).Visible = True
         b2.lblMQM(0).Visible = True
         b2.lblMTm(0).Visible = True
                  'b2.lblMQM(0).Caption = adoMM.Fields("QLabel").Value
            If adoMM.Fields("xf").Value = True Then
                b2.cmdMQm(0).Caption = adoMM.Fields("Qren").Value
                b2.lblMTm(0).Caption = adoMM.Fields("QRQ").Value
             End If
         b2.cmdMQm(0).Tag = adoMM.Fields("zid").Value
         adoMM.MoveNext
         For oo = 1 To adoMM.RecordCount - 1
           'b2.lblMQM(oo).Caption = adoMM.Fields("QLabel").Value
            If adoMM.Fields("xf").Value = True Then
                b2.cmdMQm(oo).Caption = adoMM.Fields("Qren").Value
                b2.lblMTm(oo).Caption = adoMM.Fields("QRQ").Value
           End If
           b2.cmdMQm(oo).Tag = adoMM.Fields("zid").Value
           adoMM.MoveNext
        Next

     End If

End Sub

Private Sub cmdAdd_Click()
Dim tt As String
If txt1.Text = "" Or txt2.Text = "" Or Val(txt3.Text) = 0 Or Val(lblKid.Caption) = 0 Then
    Exit Sub
End If


timZm = 2 '计划编辑
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "员工表1"
    mod1.cmd.Parameters("@NBLX") = "计划编辑"
    mod1.cmd.Parameters("@bh") = lblGid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblKid.Caption
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
    mod1.cmd.Parameters("@mlt1") = txt1.Text
    mod1.cmd.Parameters("@mlt2") = txt2.Text
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt3.Text)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 1 '添加
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
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = dtp5.Value
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV
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

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.TBa.Buttons(7).Value = tbrUnpressed
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub

Private Sub cmdBB_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
Set bView.adoBview = New ADODB.Recordset
tt = "select 部门,姓名,专项工作内容,完成期限,完成情况 from bview where uid='" & mod1.DHid & "'"
Set bView.adoBview = New ADODB.Recordset
bView.adoBview.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If bView.adoBview.RecordCount = 0 Then
    Set bView.dtgBB.DataSource = bView.adoBview
    bView.dtgBB.Rows = 2
    bView.dtgBB.FixedRows = 0
    bView.dtgBB.FixedRows = 1
Else
    bView.dtgBB.FixedRows = 1
    Set bView.dtgBB.DataSource = bView.adoBview
    bView.dtgBB.FixedRows = 0
    bView.dtgBB.MergeCol(1) = True
    bView.dtgBB.MergeCol(2) = True
    bView.dtgBB.MergeCol(5) = True
    bView.dtgBB.MergeCells = 3
    bView.dtgBB.FixedRows = 1
End If


bView.Show
Me.Enabled = False
bView.lblFw.Caption = mod1.DName
bView.lblFw.ToolTipText = mod1.DHid
bView.txtM.Value = txtM.Value
End Sub

Private Sub cmdDao_Click()
Dim tt As String
Dim ii As Integer
ii = MsgBox("是否导入上月任务?", vbQuestion + vbYesNo, "询问")
If ii = vbNo Then
    Exit Sub
End If
timZm = 2 '计划编辑
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "员工表1"
    mod1.cmd.Parameters("@NBLX") = "计划编辑"
    mod1.cmd.Parameters("@bh") = lblGid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblKid.Caption
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
    mod1.cmd.Parameters("@mlt1") = txt1.Text
    mod1.cmd.Parameters("@mlt2") = txt2.Text
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt3.Text)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 3 '删除
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
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = dtp5.Value
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV
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

Private Sub cmdDe_Click()
Dim tt As String
Dim ii As Integer
ii = MsgBox("是否删除此记录?", vbQuestion + vbYesNo, "询问")
If ii = vbNo Then
    Exit Sub
End If


timZm = 2 '计划编辑
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "员工表1"
    mod1.cmd.Parameters("@NBLX") = "计划编辑"
    mod1.cmd.Parameters("@bh") = lblGid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblKid.Caption
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
    mod1.cmd.Parameters("@mlt1") = txt1.Text
    mod1.cmd.Parameters("@mlt2") = txt2.Text
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt3.Text)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 3 '删除
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
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = dtp5.Value
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV
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

If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
timZm = 5 '签字
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "员工表1"
    mod1.cmd.Parameters("@NBLX") = "签字"
    mod1.cmd.Parameters("@bh") = lblKid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtYwy.Text
    mod1.cmd.Parameters("@mt2") = txtYwy.ToolTipText
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
    mod1.cmd.Parameters("@mt20") = lblMQM(Val(lblLc.Caption) - 1).Caption
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
    mod1.cmd.Parameters("@mm2") = Val(lblFwid.Caption)
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
    mod1.cmd.Parameters("@md1") = Null
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV
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

Private Sub cmdGx_Click()
Dim tt As String
If txt1.Text = "" Or txt2.Text = "" Or Val(txt3.Text) = 0 Or Val(lblKid.Caption) = 0 Then
    Exit Sub
End If


timZm = 2 '计划编辑
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "员工表1"
    mod1.cmd.Parameters("@NBLX") = "计划编辑"
    mod1.cmd.Parameters("@bh") = lblGid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblKid.Caption
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
    mod1.cmd.Parameters("@mlt1") = txt1.Text
    mod1.cmd.Parameters("@mlt2") = txt2.Text
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt3.Text)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mm10") = 2 '更新
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
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = dtp5.Value
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV
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

Private Sub cmdL_Click()
Dim DD As Date
DD = txtM.Value
txtM.Value = DateSerial(Year(txtM.Value), Month(txtM.Value) - 1, Day(txtM.Value))
txtGzgy.Text = ""
lblKid.Caption = ""
cmdMQm(0).Caption = ""
cmdMQm(1).Caption = ""
lblMTm(0).Caption = ""
lblMTm(1).Caption = ""
Set dtgB1.DataSource = Nothing
b2.txtA.Text = ""
Call KPIBound(txtYwy.Text, txtYwy.ToolTipText, txtM.Value)
If BF = False Then
    txtM.Value = DD
End If
frmED.Visible = False
txtGzgy.Locked = True
End Sub

Private Sub cmdMod_Click()
If mod1.DHid = lblLcUid.Caption And Val(lblKid.Caption) > 0 Then
    frmED.Visible = True
    txtGzgy.Locked = False
    cmdSave.Enabled = True
End If
End Sub

Private Sub cmdMQm_Click(Index As Integer)
Dim QZ As Integer
Dim oo As Integer
On Error Resume Next
If Me.Visible = False Then Exit Sub
'先检测权重是否超出100%
QZ = 0
dtgB1.Row = 1
dtgB1.Col = 4
'QZ = Val(dtgB1.Text)
'dtgB1.Row = dtgB1.Row + 1
'Do While Not dtgB1.Row >= dtgB1.Rows
'    QZ = QZ + Val(dtgB1.Text)
'    dtgB1.Row = dtgB1.Row + 1
'Loop
If Trim(lblLcUid.Caption) <> mod1.DHid Then
    MsgBox "此处应由" & lblLcUid.ToolTipText & "签字! 请您不要再点"
    Exit Sub
End If
For oo = 1 To dtgB1.Rows - 1
    QZ = QZ + Val(dtgB1.Text)
    dtgB1.Row = dtgB1.Row + 1
Next
If QZ <> 100 Then
    MsgBox ("重要性基数没有正确设置!")
    Exit Sub
End If
QZ = 0
QZ = Val(b2.txtC2.Text) + Val(b2.txtD2.Text) + Val(b2.txtE2.Text) + Val(b2.txtF2.Text) + Val(b2.txtG2.Text) + Val(b2.txtH2.Text) + Val(b2.txtI2.Text)
If QZ <> 20 Then
    MsgBox ("权重没有正确设置!")
    b2.Show
    b1.Visible = False
    Exit Sub
End If
If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If
If Index + 1 <> lblLc.Caption Then '不能在不相干的位置上乱点
    Exit Sub
End If




If Val(lblLc.Caption) = 1 And lblLcUid.Caption = mod1.DHid Then
    If b2.txtC1.Text = "" Or b2.txtD1.Text = "" Or b2.txtE1.Text = "" Or b2.txtF1.Text = "" Or b2.txtG1.Text = "" Or b2.txtH1.Text = "" Or b2.txtI1.Text = "" Or _
       Val(b2.txtC2.Text) = 0 Or Val(b2.txtD2.Text) = 0 Or Val(b2.txtE2.Text) = 0 Or Val(b2.txtF2.Text) = 0 Or Val(b2.txtG2.Text) = 0 Or _
         Val(b2.txtH2.Text) = 0 Or Val(b2.txtI2.Text) = 0 Then
        MsgBox "请填写相应的考核信息！"
        b2.Visible = True
        b2.ZOrder 0
        b1.Visible = False
        Exit Sub
    End If
End If

If Index = 0 Then '初次只能签字，不能驳回。
    optT2.Enabled = False
Else
    optT2.Enabled = True
End If
OptT1.Value = True

frmQm.Visible = True

End Sub

Private Sub cmdNew_Click()
Dim tt As String
On Error Resume Next

timZm = 1 '员工表1新建
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "员工表1"
    mod1.cmd.Parameters("@NBLX") = "新建"
    mod1.cmd.Parameters("@bh") = ""
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtYwy.Text
    mod1.cmd.Parameters("@mt2") = txtYwy.ToolTipText
    mod1.cmd.Parameters("@mt3") = txtBm.Text
    mod1.cmd.Parameters("@mt4") = txtZw.Text
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
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = 0
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
    mod1.cmd.Parameters("@mm16") = 0
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = txtM.Value
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV
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

Private Sub cmdQing_Click()
txt1.Text = ""
txt2.Text = ""
txt3.Text = ""
End Sub

Private Sub cmdR_Click()
Dim DD As Date
DD = txtM.Value
txtM.Value = DateSerial(Year(txtM.Value), Month(txtM.Value) + 1, Day(txtM.Value))
txtGzgy.Text = ""
lblKid.Caption = ""
cmdMQm(0).Caption = ""
cmdMQm(1).Caption = ""
lblMTm(0).Caption = ""
lblMTm(1).Caption = ""
Set dtgB1.DataSource = Nothing
b2.txtA.Text = ""
Call KPIBound(txtYwy.Text, txtYwy.ToolTipText, txtM.Value)
If BF = False Then
    txtM.Value = DD
End If
frmED.Visible = False
txtGzgy.Locked = True
End Sub


Private Sub cmdSave_Click()
Dim tt As String
On Error Resume Next


timZm = 3 '员工表1保存
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "员工表1"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = lblKid.Caption
'If Val(lblLc.Caption) = 0 Then
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtYwy.Text
    mod1.cmd.Parameters("@mt2") = txtYwy.ToolTipText
    mod1.cmd.Parameters("@mt3") = txtBm.Text
    mod1.cmd.Parameters("@mt4") = txtZw.Text
    mod1.cmd.Parameters("@mt5") = txtGzgy.Text
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
    mod1.cmd.Parameters("@mlt2") = ""
    mod1.cmd.Parameters("@mlt3") = ""
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = 0
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
    mod1.cmd.Parameters("@mm16") = 0
    mod1.cmd.Parameters("@mm17") = 0
    mod1.cmd.Parameters("@mm18") = 0
    mod1.cmd.Parameters("@mm19") = 0
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@mb2") = 0
    mod1.cmd.Parameters("@mb3") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = txtM.Value
    mod1.cmd.Parameters("@md2") = Null
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV
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

Private Sub cmdView_Click()
    Set Ren.XForm = New b1
    Call mod1.RenXz("b1", Me, 0)
End Sub

Private Sub cmdZuan_Click()
If b1.Visible = True Then
    b2.Visible = True
    b1.Visible = False
ElseIf b2.Visible = True Then
    b3.Visible = True
    b2.Visible = False
ElseIf b3.Visible = True Then
    b1.Visible = True
    b3.Visible = False
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dtgB1_Click()
On Error Resume Next
If dtgB1.Row = 0 Then Exit Sub
dtgB1.Col = 1
txt1.Text = dtgB1.Text
dtgB1.Col = 2
txt2.Text = dtgB1.Text
dtgB1.Col = 3
dtp5.Value = dtgB1.Text
dtgB1.Col = 4
txt3.Text = dtgB1.Text
dtgB1.Col = 5
lblGid.Caption = Val(dtgB1.Text)
frmQm.Visible = False
End Sub

Private Sub dtgB1_RowColChange()
On Error Resume Next
If dtgB1.Row = 0 Then Exit Sub
dtgB1.Col = 1
txt1.Text = dtgB1.Text
dtgB1.Col = 2
txt2.Text = dtgB1.Text
dtgB1.Col = 3
dtp5.Value = dtgB1.Text
dtgB1.Col = 4
txt3.Text = dtgB1.Text
dtgB1.Col = 5
lblGid.Caption = Val(dtgB1.Text)
End Sub

Private Sub Form_Click()
frmQm.Visible = False
lblTX.Visible = False
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
dtgB1.ColWidth(0) = 300
dtgB1.ColWidth(1) = 7320
dtgB1.ColWidth(2) = 3920
dtgB1.ColWidth(3) = 1020
dtgB1.ColWidth(4) = 1000
dtgB1.ColWidth(5) = 0
dtp5.Value = mod1.DQda
txtM.Value = mod1.DQda
frmQm.Left = 450
frmQm.Top = 7440
frmQm.Visible = False
End Sub


Public Sub KPIBound1(Kid As Long, Lei As String)
Dim tt As String
Dim oo As Integer
On Error Resume Next
Dialog.OBF = False

tt = "select gzgy,kid,lc,lcuid,fwid,uid,bm,zw from b1 where kid=" & Kid
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly
If IsNull(mod1.HTP.RecordCount) = True Then
    Dialog.Enabled = True
    MsgBox "读取数据出错！"
    Exit Sub
End If
txtGzgy.Text = mod1.HTP.Fields("gzgy").Value
lblKid.Caption = mod1.HTP.Fields("kid").Value
lblLc.Caption = mod1.HTP.Fields("lc").Value
lblLcUid.Caption = mod1.HTP.Fields("lcuid").Value
lblFwid.Caption = mod1.HTP.Fields("fwid").Value
txtBm.Text = mod1.HTP.Fields("bm").Value
txtZw.Text = mod1.HTP.Fields("zw").Value
If Val(lblKid.Caption) > 0 Then
    cmdNew.Visible = False
    frmED.Visible = False
Else
    cmdNew.Visible = True
End If

txtYwy.ToolTipText = mod1.HTP.Fields("uid").Value
tt = "select username from worker where userid='" & Trim(mod1.HTP.Fields("uid").Value) & "'"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly
If mod1.HTP.RecordCount = 0 Then
    Dialog.Enabled = True
    MsgBox "读取数据出错！"
    Exit Sub
End If
txtYwy.Text = mod1.HTP.Fields("username").Value

'tt = "select username from worker where userid='" & lblLcUid.Caption & "'"
'Set mod1.HTP = New ADODB.Recordset
'mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly
'If mod1.HTP.RecordCount = 0 Then
'    MsgBox "读取数据出错！"
'    Exit Sub
'End If
'lblLcUid.ToolTipText = mod1.HTP.Fields("username").Value

tt = "select gnr as 月度工作计划内容,gzmb as 工作目标,wrq as 完成时间,zj as 重要性基数,gid from b11 where kid=" & Val(lblKid.Caption)
Set M1 = New ADODB.Recordset
M1.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(M1.RecordCount) = True Then
    Dialog.Enabled = True
    MsgBox "读取数据出错！"
    Exit Sub
End If
Set dtgB1.DataSource = M1
If M1.RecordCount = 0 Then
    dtgB1.Rows = 2
    dtgB1.FixedRows = 0
    dtgB1.FixedRows = 1
End If
Call OAn
'打开表2
b2.txtYwy.Text = txtYwy.Text
b2.txtYwy.ToolTipText = txtYwy.ToolTipText
b2.txtM.Value = txtM.Value
b2.txtBm.Text = txtBm.Text
M1.MoveFirst
oo = 1
Do While Not M1.EOF
    b2.txtA.Text = b2.txtA.Text & oo & "." & M1.Fields("月度工作计划内容").Value & Chr(13) & " "
    oo = oo + 1
    M1.MoveNext
Loop
tt = "select * from b2 where kid=" & Val(lblKid.Caption)
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly
If IsNull(mod1.HTP.RecordCount) = True Then
    MsgBox "读取数据出错！"
    Exit Sub
End If
b2.txtA1.Text = mod1.HTP.Fields("a1").Value
b2.txtA2.Text = mod1.HTP.Fields("a2").Value
b2.txtA3.Text = mod1.HTP.Fields("a3").Value
b2.txtA4.Text = mod1.HTP.Fields("a4").Value
b2.txtA5.Text = mod1.HTP.Fields("a5").Value
b2.txtB1.Text = mod1.HTP.Fields("b1").Value
b2.txtB2.Text = mod1.HTP.Fields("b2").Value
b2.txtB3.Text = mod1.HTP.Fields("b3").Value
b2.txtB4.Text = mod1.HTP.Fields("b4").Value
b2.txtB5.Text = mod1.HTP.Fields("b5").Value
b2.txtC1.Text = mod1.HTP.Fields("c1").Value
b2.txtC2.Text = mod1.HTP.Fields("c2").Value
b2.txtC3.Text = mod1.HTP.Fields("c3").Value
b2.txtC4.Text = mod1.HTP.Fields("c4").Value
b2.txtC5.Text = mod1.HTP.Fields("c5").Value
b2.txtD1.Text = mod1.HTP.Fields("d1").Value
b2.txtD2.Text = mod1.HTP.Fields("d2").Value
b2.txtD3.Text = mod1.HTP.Fields("d3").Value
b2.txtD4.Text = mod1.HTP.Fields("d4").Value
b2.txtD5.Text = mod1.HTP.Fields("d5").Value
b2.txtE1.Text = mod1.HTP.Fields("e1").Value
b2.txtE2.Text = mod1.HTP.Fields("e2").Value
b2.txtE3.Text = mod1.HTP.Fields("e3").Value
b2.txtE4.Text = mod1.HTP.Fields("e4").Value
b2.txtE5.Text = mod1.HTP.Fields("e5").Value
b2.txtF1.Text = mod1.HTP.Fields("f1").Value
b2.txtF2.Text = mod1.HTP.Fields("f2").Value
b2.txtF3.Text = mod1.HTP.Fields("f3").Value
b2.txtF4.Text = mod1.HTP.Fields("f4").Value
b2.txtF5.Text = mod1.HTP.Fields("f5").Value
b2.txtG1.Text = mod1.HTP.Fields("g1").Value
b2.txtG2.Text = mod1.HTP.Fields("g2").Value
b2.txtG3.Text = mod1.HTP.Fields("g3").Value
b2.txtG4.Text = mod1.HTP.Fields("g4").Value
b2.txtG5.Text = mod1.HTP.Fields("g5").Value
b2.txtH1.Text = mod1.HTP.Fields("h1").Value
b2.txtH2.Text = mod1.HTP.Fields("h2").Value
b2.txtH3.Text = mod1.HTP.Fields("h3").Value
b2.txtH4.Text = mod1.HTP.Fields("h4").Value
b2.txtH5.Text = mod1.HTP.Fields("h5").Value
b2.txtI1.Text = mod1.HTP.Fields("i1").Value
b2.txtI2.Text = mod1.HTP.Fields("i2").Value
b2.txtI3.Text = mod1.HTP.Fields("i3").Value
b2.txtI4.Text = mod1.HTP.Fields("i4").Value
b2.txtI5.Text = mod1.HTP.Fields("i5").Value
b2.txtJ1.Text = mod1.HTP.Fields("j1").Value
b2.txtJ2.Text = mod1.HTP.Fields("j2").Value
b2.txtJ3.Text = mod1.HTP.Fields("j3").Value
b2.txtJ4.Text = mod1.HTP.Fields("j4").Value
b2.txtJ5.Text = mod1.HTP.Fields("j5").Value
b2.lblZF.Caption = mod1.HTP.Fields("zf").Value
b2.txtZjp.Text = mod1.HTP.Fields("zjp").Value
b2.txtBmp.Text = mod1.HTP.Fields("bmp").Value
b2.lblKid.Caption = lblKid.Caption
b2.lblLc.Caption = mod1.HTP.Fields("lc").Value
b2.lblLcUid.Caption = mod1.HTP.Fields("lcuid").Value
b2.lblFwid.Caption = mod1.HTP.Fields("fwid").Value
b2.txtC2.Text = Val(b2.txtC2.Text) & "%"
b2.txtD2.Text = Val(b2.txtD2.Text) & "%"
b2.txtE2.Text = Val(b2.txtE2.Text) & "%"
b2.txtF2.Text = Val(b2.txtF2.Text) & "%"
b2.txtG2.Text = Val(b2.txtG2.Text) & "%"
b2.txtH2.Text = Val(b2.txtH2.Text) & "%"
b2.txtI2.Text = Val(b2.txtI2.Text) & "%"
b2.txtJ2.Text = Val(b2.txtJ2.Text) & "%"
b2.lblZ3.Caption = Val(b2.txtA3.Text) * Val(b2.txtA2.Text) / 100 + Val(b2.txtB3.Text) * Val(b2.txtB2.Text) / 100 + Val(b2.txtC3.Text) * Val(b2.txtC2.Text) / 100 + _
Val(b2.txtD3.Text) * Val(b2.txtD2.Text) / 100 + Val(b2.txtE3.Text) * Val(b2.txtE2.Text) / 100 + Val(b2.txtF3.Text) * Val(b2.txtF2.Text) / 100 + Val(b2.txtG3.Text) * Val(b2.txtG2.Text) / 100 + _
Val(b2.txtH3.Text) * Val(b2.txtH2.Text) / 100 + Val(b2.txtI3.Text) * Val(b2.txtI2.Text) / 100
b2.lblZ4.Caption = Val(b2.txtA4.Text) * Val(b2.txtA2.Text) / 100 + Val(b2.txtB4.Text) * Val(b2.txtB2.Text) / 100 + Val(b2.txtC4.Text) * Val(b2.txtC2.Text) / 100 + _
Val(b2.txtD4.Text) * Val(b2.txtD2.Text) / 100 + Val(b2.txtE4.Text) * Val(b2.txtE2.Text) / 100 + Val(b2.txtF4.Text) * Val(b2.txtF2.Text) / 100 + Val(b2.txtG4.Text) * Val(b2.txtG2.Text) / 100 + _
Val(b2.txtH4.Text) * Val(b2.txtH2.Text) / 100 + Val(b2.txtI4.Text) * Val(b2.txtI2.Text) / 100
b2.lblZ5.Caption = Val(b2.txtA5.Text) * Val(b2.txtA2.Text) / 100 + Val(b2.txtB5.Text) * Val(b2.txtB2.Text) / 100 + Val(b2.txtC5.Text) * Val(b2.txtC2.Text) / 100 + _
Val(b2.txtD5.Text) * Val(b2.txtD2.Text) / 100 + Val(b2.txtE5.Text) * Val(b2.txtE2.Text) / 100 + Val(b2.txtF5.Text) * Val(b2.txtF2.Text) / 100 + Val(b2.txtG5.Text) * Val(b2.txtG2.Text) / 100 + _
Val(b2.txtH5.Text) * Val(b2.txtH2.Text) / 100 + Val(b2.txtI5.Text) * Val(b2.txtI2.Text) / 100
Dialog.OBF = True
If Lei = "月度工作计划" Then
    b1.Show
ElseIf Lei = "员工月度考核表" Then
    b2.Show
End If
End Sub

Public Sub KPIBound(Ywy As String, Uid As String, DD As Date)
Dim tt As String
Dim oo As Integer
On Error Resume Next
BF = False
txtYwy.Text = Ywy
txtYwy.ToolTipText = Uid
'txtM.ToolTipText = DD
'txtM.Text = Year(DD) & "年" & Month(DD) & "月"
tt = "select bm,userzw,ggl from worker where userid='" & Uid & "'"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly
If mod1.HTP.RecordCount = 0 Then
    MsgBox "读取数据出错！"
    Exit Sub
End If
txtBm.Text = mod1.HTP.Fields("bm").Value
txtZw.Text = mod1.HTP.Fields("userzw").Value
yGGl = mod1.HTP.Fields("ggl").Value
tt = "select gzgy,kid,lc,lcuid,fwid from b1 where uid='" & Uid & "' and year(yf)=" & Year(txtM.Value) & " and month(yf)=" & Month(txtM.Value)
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly
If IsNull(mod1.HTP.RecordCount) = True Then
    MsgBox "读取数据出错！"
    Exit Sub
End If
txtGzgy.Text = mod1.HTP.Fields("gzgy").Value
lblKid.Caption = mod1.HTP.Fields("kid").Value
lblLc.Caption = mod1.HTP.Fields("lc").Value
lblLcUid.Caption = mod1.HTP.Fields("lcuid").Value
lblFwid.Caption = mod1.HTP.Fields("fwid").Value
If Val(lblKid.Caption) > 0 Or mod1.DHid <> yGGl Then
    cmdNew.Visible = False
    frmED.Visible = False
Else
    cmdNew.Visible = True
End If
If mod1.HTP.RecordCount > 0 Then
    tt = "select username from worker where userid='" & lblLcUid.Caption & "'"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly
    If IsNull(mod1.HTP.RecordCount) = True Then
        MsgBox "读取数据出错！"
        Exit Sub
    End If
    lblLcUid.ToolTipText = mod1.HTP.Fields("username").Value
End If
tt = "select gnr as 月度工作计划内容,gzmb as 工作目标,wrq as 完成时间,zj as 重要性基数,gid from b11 where kid=" & Val(lblKid.Caption)
Set M1 = New ADODB.Recordset
M1.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(M1.RecordCount) = True Then
    MsgBox "读取数据出错！"
    Exit Sub
End If
Set dtgB1.DataSource = M1
If M1.RecordCount = 0 Then
    dtgB1.Rows = 2
    dtgB1.FixedRows = 0
    dtgB1.FixedRows = 1
End If
Call OAn

'打开表2
b2.txtYwy.Text = txtYwy.Text
b2.txtYwy.ToolTipText = txtYwy.ToolTipText
b2.txtM.Value = txtM.Value
b2.txtBm.Text = txtBm.Text
M1.MoveFirst
oo = 1
Do While Not M1.EOF
    b2.txtA.Text = b2.txtA.Text & oo & "." & M1.Fields("月度工作计划内容").Value & ";" & Chr(13) & " "
    oo = oo + 1
    M1.MoveNext
Loop
tt = "select * from b2 where kid=" & Val(lblKid.Caption)
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly
If IsNull(mod1.HTP.RecordCount) = True Then
    MsgBox "读取数据出错！"
    Exit Sub
End If
b2.txtA1.Text = mod1.HTP.Fields("a1").Value
b2.txtA2.Text = mod1.HTP.Fields("a2").Value
b2.txtA3.Text = mod1.HTP.Fields("a3").Value
b2.txtA4.Text = mod1.HTP.Fields("a4").Value
b2.txtA5.Text = mod1.HTP.Fields("a5").Value
b2.txtB1.Text = mod1.HTP.Fields("b1").Value
b2.txtB2.Text = mod1.HTP.Fields("b2").Value
b2.txtB3.Text = mod1.HTP.Fields("b3").Value
b2.txtB4.Text = mod1.HTP.Fields("b4").Value
b2.txtB5.Text = mod1.HTP.Fields("b5").Value
b2.txtC1.Text = mod1.HTP.Fields("c1").Value
b2.txtC2.Text = mod1.HTP.Fields("c2").Value
b2.txtC3.Text = mod1.HTP.Fields("c3").Value
b2.txtC4.Text = mod1.HTP.Fields("c4").Value
b2.txtC5.Text = mod1.HTP.Fields("c5").Value
b2.txtD1.Text = mod1.HTP.Fields("d1").Value
b2.txtD2.Text = mod1.HTP.Fields("d2").Value
b2.txtD3.Text = mod1.HTP.Fields("d3").Value
b2.txtD4.Text = mod1.HTP.Fields("d4").Value
b2.txtD5.Text = mod1.HTP.Fields("d5").Value
b2.txtE1.Text = mod1.HTP.Fields("e1").Value
b2.txtE2.Text = mod1.HTP.Fields("e2").Value
b2.txtE3.Text = mod1.HTP.Fields("e3").Value
b2.txtE4.Text = mod1.HTP.Fields("e4").Value
b2.txtE5.Text = mod1.HTP.Fields("e5").Value
b2.txtF1.Text = mod1.HTP.Fields("f1").Value
b2.txtF2.Text = mod1.HTP.Fields("f2").Value
b2.txtF3.Text = mod1.HTP.Fields("f3").Value
b2.txtF4.Text = mod1.HTP.Fields("f4").Value
b2.txtF5.Text = mod1.HTP.Fields("f5").Value
b2.txtG1.Text = mod1.HTP.Fields("g1").Value
b2.txtG2.Text = mod1.HTP.Fields("g2").Value
b2.txtG3.Text = mod1.HTP.Fields("g3").Value
b2.txtG4.Text = mod1.HTP.Fields("g4").Value
b2.txtG5.Text = mod1.HTP.Fields("g5").Value
b2.txtH1.Text = mod1.HTP.Fields("h1").Value
b2.txtH2.Text = mod1.HTP.Fields("h2").Value
b2.txtH3.Text = mod1.HTP.Fields("h3").Value
b2.txtH4.Text = mod1.HTP.Fields("h4").Value
b2.txtH5.Text = mod1.HTP.Fields("h5").Value
b2.txtI1.Text = mod1.HTP.Fields("i1").Value
b2.txtI2.Text = mod1.HTP.Fields("i2").Value
b2.txtI3.Text = mod1.HTP.Fields("i3").Value
b2.txtI4.Text = mod1.HTP.Fields("i4").Value
b2.txtI5.Text = mod1.HTP.Fields("i5").Value
b2.txtJ1.Text = mod1.HTP.Fields("j1").Value
b2.txtJ2.Text = mod1.HTP.Fields("j2").Value
b2.txtJ3.Text = mod1.HTP.Fields("j3").Value
b2.txtJ4.Text = mod1.HTP.Fields("j4").Value
b2.txtJ5.Text = mod1.HTP.Fields("j5").Value

b2.lblZF.Caption = mod1.HTP.Fields("zf").Value

b2.txtZjp.Text = mod1.HTP.Fields("zjp").Value
b2.txtBmp.Text = mod1.HTP.Fields("bmp").Value
b2.lblKid.Caption = lblKid.Caption
b2.lblFwid.Caption = mod1.HTP.Fields("fwid").Value
b2.lblLc.Caption = mod1.HTP.Fields("lc").Value
b2.lblLcUid.Caption = mod1.HTP.Fields("lcuid").Value
b2.txtC2.Text = Val(b2.txtC2.Text) & "%"
b2.txtD2.Text = Val(b2.txtD2.Text) & "%"
b2.txtE2.Text = Val(b2.txtE2.Text) & "%"
b2.txtF2.Text = Val(b2.txtF2.Text) & "%"
b2.txtG2.Text = Val(b2.txtG2.Text) & "%"
b2.txtH2.Text = Val(b2.txtH2.Text) & "%"
b2.txtI2.Text = Val(b2.txtI2.Text) & "%"
b2.txtJ2.Text = Val(b2.txtJ2.Text) & "%"
b2.lblZ3.Caption = Val(b2.txtA3.Text) * Val(b2.txtA2.Text) / 100 + Val(b2.txtB3.Text) * Val(b2.txtB2.Text) / 100 + Val(b2.txtC3.Text) * Val(b2.txtC2.Text) / 100 + _
Val(b2.txtD3.Text) * Val(b2.txtD2.Text) / 100 + Val(b2.txtE3.Text) * Val(b2.txtE2.Text) / 100 + Val(b2.txtF3.Text) * Val(b2.txtF2.Text) / 100 + Val(b2.txtG3.Text) * Val(b2.txtG2.Text) / 100 + _
Val(b2.txtH3.Text) * Val(b2.txtH2.Text) / 100 + Val(b2.txtI3.Text) * Val(b2.txtI2.Text) / 100
b2.lblZ4.Caption = Val(b2.txtA4.Text) * Val(b2.txtA2.Text) / 100 + Val(b2.txtB4.Text) * Val(b2.txtB2.Text) / 100 + Val(b2.txtC4.Text) * Val(b2.txtC2.Text) / 100 + _
Val(b2.txtD4.Text) * Val(b2.txtD2.Text) / 100 + Val(b2.txtE4.Text) * Val(b2.txtE2.Text) / 100 + Val(b2.txtF4.Text) * Val(b2.txtF2.Text) / 100 + Val(b2.txtG4.Text) * Val(b2.txtG2.Text) / 100 + _
Val(b2.txtH4.Text) * Val(b2.txtH2.Text) / 100 + Val(b2.txtI4.Text) * Val(b2.txtI2.Text) / 100
b2.lblZ5.Caption = Val(b2.txtA5.Text) * Val(b2.txtA2.Text) / 100 + Val(b2.txtB5.Text) * Val(b2.txtB2.Text) / 100 + Val(b2.txtC5.Text) * Val(b2.txtC2.Text) / 100 + _
Val(b2.txtD5.Text) * Val(b2.txtD2.Text) / 100 + Val(b2.txtE5.Text) * Val(b2.txtE2.Text) / 100 + Val(b2.txtF5.Text) * Val(b2.txtF2.Text) / 100 + Val(b2.txtG5.Text) * Val(b2.txtG2.Text) / 100 + _
Val(b2.txtH5.Text) * Val(b2.txtH2.Text) / 100 + Val(b2.txtI5.Text) * Val(b2.txtI2.Text) / 100


BF = True
b1.Show

End Sub
Public Sub KPIQing()
txtYwy.Text = ""
txtYwy.ToolTipText = ""
txtBm.Text = ""
txtZw.Text = ""
'txtM.Text = ""
'txtM.ToolTipText = ""
txtGzgy.Text = ""
txtGzgy.Locked = True
Set dtgB1.DataSource = Nothing
lblKid.Caption = ""
lblGid.Caption = ""
cmdNew.Visible = False
frmED.Visible = False
lblLcUid.Caption = ""
lblLc.Caption = ""
lblFwid.Caption = ""
cmdMQm(0).Caption = ""
cmdMQm(1).Caption = ""
lblMTm(0).Caption = ""
lblMTm(1).Caption = ""
b2.txtYwy.Text = ""
b2.txtYwy.ToolTipText = ""
b2.txtBm.Text = ""
b2.txtA.Text = ""
b2.txtA1.Text = ""
'b2.txtA2.Text = ""
b2.txtA3.Text = ""
b2.txtA4.Text = ""
b2.txtA5.Text = 100
b2.txtB1.Text = ""
'b2.txtB2.Text = ""
b2.txtB3.Text = ""
b2.txtB4.Text = ""
b2.txtB5.Text = 100
b2.txtC1.Text = ""
b2.txtC2.Text = ""
b2.txtC3.Text = ""
b2.txtC4.Text = ""
b2.txtC5.Text = 100
b2.txtD1.Text = ""
b2.txtD2.Text = ""
b2.txtD3.Text = ""
b2.txtD4.Text = ""
b2.txtD5.Text = 100
b2.txtE1.Text = ""
b2.txtE2.Text = ""
b2.txtE3.Text = ""
b2.txtE4.Text = ""
b2.txtE5.Text = 100
b2.txtF1.Text = ""
b2.txtF2.Text = ""
b2.txtF3.Text = ""
b2.txtF4.Text = ""
b2.txtF5.Text = 100
b2.txtG1.Text = ""
b2.txtG2.Text = ""
b2.txtG3.Text = ""
b2.txtG4.Text = ""
b2.txtG5.Text = 100
b2.txtH1.Text = ""
b2.txtH2.Text = ""
b2.txtH3.Text = ""
b2.txtH4.Text = ""
b2.txtH5.Text = 100
b2.txtI1.Text = ""
b2.txtI2.Text = ""
b2.txtI3.Text = ""
b2.txtI4.Text = ""
b2.txtI5.Text = 100
b2.txtJ1.Text = ""
b2.txtJ2.Text = ""
b2.txtJ3.Text = ""
b2.txtJ4.Text = ""
b2.txtJ5.Text = 100
b2.lblZ3.Caption = 0
b2.lblZ4.Caption = 0
b2.lblZ5.Caption = 100
b2.lblZF.Caption = ""
b2.txtZjp.Text = ""
b2.txtBmp.Text = ""
b2.lblLc.Caption = ""
b2.lblLcUid.Caption = ""
b2.lblFwid.Caption = ""
b2.lblKid.Caption = ""
cmdSave.Enabled = False
b2.cmdSave.Enabled = False
Call b2.b2Locked
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmZu.TBa.Buttons(7).Value = tbrUnpressed
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub

Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '如果为添加合同评审
    frmED.Visible = True
    txtGzgy.Locked = False
    cmdNew.Visible = False
    cmdSave.Enabled = True
    lblLc.Caption = 1
    lblLcUid.Caption = mod1.DHid
    b2.lblLc.Caption = 0
    b2.lblLcUid.Caption = mod1.DHid
ElseIf timZm = 2 Then '计划编辑
    txt1.Text = ""
    txt2.Text = ""
    txt3.Text = ""
    lblGid.Caption = ""
    tt = "select gnr as 月度工作计划内容,gzmb as 工作目标,wrq as 完成时间,zj as 重要性基数,gid from b11 where kid=" & Val(lblKid.Caption)
    Set M1 = New ADODB.Recordset
    M1.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(M1.RecordCount) = True Then
    MsgBox "读取数据出错！"
    Exit Sub
End If
    Set dtgB1.DataSource = M1
    If M1.RecordCount = 0 Then
        dtgB1.Rows = 2
        dtgB1.FixedRows = 0
        dtgB1.FixedRows = 1
    End If
ElseIf timZm = 3 Then
    txtGzgy.Locked = True
    cmdSave.Enabled = False
    frmED.Visible = False
ElseIf timZm = 5 Then '签字
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
    Call mod1.refEnvent
    b2.lblLc.Caption = 1
    b2.lblLcUid.Caption = txtYwy.ToolTipText
    If Val(lblLc.Caption) > 2 Then
        MsgBox "请慎重对待你所承诺的工作！"
    ElseIf Val(lblLc.Caption) = 2 Then
    
    End If
End If
timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2 from ml where zid=" & mod1.Zid
Set mod1.WP = New ADODB.Recordset
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        lblFwid.Caption = mod1.WP.Fields("mm1").Value
        lblKid.Caption = mod1.WP.Fields("bh").Value
        b2.lblKid.Caption = mod1.WP.Fields("bh").Value
    ElseIf timZm = 5 Then '签名
        If OptT1.Value = True Then
            cmdMQm(lblLc.Caption - 1).Caption = mod1.DName
            lblMTm(lblLc.Caption - 1).Caption = mod1.DQda
        Else
            For oo = 0 To 5
                cmdMQm(oo).Caption = ""
                lblMTm(oo).Caption = ""
            Next
        End If
        lblLc.Caption = mod1.WP.Fields("mm1").Value
        lblFwid.Caption = mod1.WP.Fields("mm2").Value
        lblLcUid.ToolTipText = mod1.WP.Fields("mt1").Value
        lblLcUid.Caption = mod1.WP.Fields("mt2").Value
        lblTX.Caption = "下一流程,将跳至" & lblMQM(Val(lblLc.Caption) - 1).Caption & ": " & lblLcUid.ToolTipText
        lblTX.Visible = True
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
        cmdNew.Enabled = False
    End If
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdNew.Enabled = False
    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


Private Sub txt1_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(txt1.Text) >= txt1.Tag Then
    MsgBox ("字数受限！")
    Exit Sub
End If
End Sub


Private Sub txt2_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(txt2.Text) >= txt2.Tag Then
    MsgBox ("字数受限！")
    Exit Sub
End If
End Sub


Private Sub txtGzgy_KeyDown(KeyCode As Integer, Shift As Integer)
If Len(txtGzgy.Text) >= txtGzgy.Tag Then
    MsgBox ("字数受限！")
    Exit Sub
End If

End Sub


