VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmHtz1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "执行状况"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   5100
      TabIndex        =   25
      Top             =   7530
      Visible         =   0   'False
      Width           =   6315
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1305
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   300
         Width           =   4965
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   28
         Top             =   510
         Width           =   705
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   27
         Top             =   870
         Width           =   675
      End
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdZX 
      BackColor       =   &H00FFC0C0&
      Caption         =   "确认执行"
      Height          =   585
      Left            =   11850
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8580
      Width           =   615
   End
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "付款审核"
      Height          =   585
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8580
      Width           =   645
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   810
      Top             =   0
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   30
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   375
      Left            =   6960
      TabIndex        =   24
      Top             =   660
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "提交"
      Height          =   585
      Left            =   13860
      Picture         =   "frmHtz1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8580
      Width           =   675
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   585
      Left            =   13200
      Picture         =   "frmHtz1.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8580
      Width           =   645
   End
   Begin VB.Frame frmEd 
      BackColor       =   &H00C0FFC0&
      Caption         =   "付款编辑"
      Height          =   1725
      Left            =   10350
      TabIndex        =   14
      Top             =   1470
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtTrader 
         BackColor       =   &H00BFFFE2&
         Height          =   270
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   270
         Width           =   3525
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "删除"
         Height          =   315
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1260
         Width           =   645
      End
      Begin VB.CommandButton cmdBG 
         BackColor       =   &H00FFFFC0&
         Caption         =   "关闭"
         Height          =   315
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1260
         Width           =   765
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFC0&
         Caption         =   "提交"
         Height          =   315
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1260
         Width           =   825
      End
      Begin VB.TextBox txtJe 
         BackColor       =   &H00BFFFE2&
         Height          =   285
         Left            =   1170
         TabIndex        =   20
         Top             =   750
         Width           =   2475
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "供应商"
         Height          =   165
         Left            =   330
         TabIndex        =   34
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "金额"
         Height          =   255
         Left            =   510
         TabIndex        =   19
         Top             =   780
         Width           =   585
      End
   End
   Begin VB.TextBox txtBz 
      BackColor       =   &H00BFFFE2&
      Height          =   1635
      Left            =   0
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   6720
      Width           =   15165
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14550
      Picture         =   "frmHtz1.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8580
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   3135
      Left            =   0
      TabIndex        =   22
      Top             =   3150
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   5530
      _Version        =   393216
      BackColor       =   15728356
      ForeColor       =   0
      Rows            =   15
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   12648447
      ForeColorFixed  =   0
      BackColorSel    =   15728356
      BackColorBkg    =   15728356
      GridColorFixed  =   12640511
      GridColorUnpopulated=   12640511
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMa 
      Height          =   2175
      Left            =   30
      TabIndex        =   36
      Top             =   930
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   12648384
      Rows            =   10
      Cols            =   8
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      BackColorUnpopulated=   8454016
      GridColorUnpopulated=   8454016
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin VB.Label lblTX 
      BackStyle       =   0  'Transparent
      Caption         =   "流程至:"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   90
      TabIndex        =   31
      Top             =   8730
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Label lblzTime 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1260
      TabIndex        =   17
      Top             =   600
      Width           =   2115
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "执行时间"
      Height          =   195
      Left            =   180
      TabIndex        =   16
      Top             =   600
      Width           =   885
   End
   Begin VB.Label lblJe 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   5160
      TabIndex        =   13
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label lbl111 
      BackStyle       =   0  'Transparent
      Caption         =   "金额(已支付)"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   3840
      TabIndex        =   12
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label lblFwid 
      Caption         =   "Label5"
      Height          =   315
      Left            =   4770
      TabIndex        =   11
      Top             =   8610
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "备注"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   6420
      Width           =   855
   End
   Begin VB.Label lblYwy 
      BackStyle       =   0  'Transparent
      Caption         =   "Label7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11520
      TabIndex        =   8
      Top             =   180
      Width           =   1065
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "业务员"
      Height          =   225
      Left            =   10620
      TabIndex        =   7
      Top             =   180
      Width           =   795
   End
   Begin VB.Label lblXz 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8700
      TabIndex        =   6
      Top             =   180
      Width           =   945
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "性质"
      Height          =   225
      Left            =   7980
      TabIndex        =   5
      Top             =   180
      Width           =   585
   End
   Begin VB.Label lblHtbh 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4920
      TabIndex        =   4
      Top             =   180
      Width           =   2925
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "合同编号"
      Height          =   225
      Left            =   3840
      TabIndex        =   3
      Top             =   180
      Width           =   795
   End
   Begin VB.Label lblXmmc 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1260
      TabIndex        =   2
      Top             =   180
      Width           =   2265
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "项目名称"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   1575
   End
End
Attribute VB_Name = "frmHtz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Fwid As Long
Public LCUid As String
Public LCRen As String
Public Hid As Long
Public Pwf As Boolean
Public Zid As Long
Public Uid As String
Dim timZm As Integer '数据提交后,由timWait执行的后续命令ID(1 添加付款 2确认执行 3付款审核 5 保存 6删除付款)
Public ZT As String '状态
Dim Lc As Integer
Public Fid As Long
Dim Amount As Single
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
dtgP.ColWidth(2) = 0
 dtgP.ColWidth(3) = 9840: dtgP.ColWidth(4) = 975
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub

Private Sub cmdAdd_Click()
Dim tt As String
Dim YYY As Long
Dim hg As Single
Dim oo As Integer
On Error Resume Next
If Me.ZT = "未执行" Then
    MsgBox "此执行单还未被确认执行,不能够生成付款!"
    Exit Sub
End If
If Val(txtJe.Text) = 0 Then
    Exit Sub
End If

dtgN.Row = dtgN.Rows - 1
dtgN.Col = 9 'pwf
If dtgN.Text = "False" And dtgN.Row > 1 Then
    MsgBox "上笔付款还未实际支付,不能够生成新付款!"
    Exit Sub
End If


timZm = 1 '添加付款
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "合同执行"
    mod1.cmd.Parameters("@NBLX") = "添加付款"
    mod1.cmd.Parameters("@bh") = Zid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtJe.Text)
    mod1.cmd.Parameters("@mm2") = txtTrader.Tag
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
        If timZm = 3 Then '保存

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

End Sub

Private Sub cmdBack_Click()
Me.Visible = False
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub

Private Sub cmdBG_Click()
frmED.Visible = False
End Sub

Private Sub cmdDel_Click()
Dim tt As String
Dim ii As Integer
Dim hg As Single
Dim oo As Integer
On Error Resume Next



dtgN.Row = dtgN.Rows - 1
dtgN.Col = 9 'pwf
If dtgN.Text = "True" Then
    MsgBox "已经支付!不能删除!"
    Exit Sub
End If
'If Me.Lc > 1 Then
dtgN.Col = 2
If dtgN.Text <> "" Then
    MsgBox "已经进入付款流程,不能删除!"
    Exit Sub
End If
ii = MsgBox("是否确认删除此笔付款?", vbQuestion + vbYesNo, "Hello")
If ii = vbNo Then Exit Sub
dtgN.Col = 10
Fid = Val(dtgN.Text)

timZm = 6 '删除付款
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "合同执行"
    mod1.cmd.Parameters("@NBLX") = "删除付款"
    mod1.cmd.Parameters("@bh") = Zid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mm1") = Fid
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
        If timZm = 3 Then '保存

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

End Sub

Private Sub cmdDing_Click()
Dim tt As String
Dim XJF As Integer
On Error Resume Next
If OptT1.Value = False And optT2.Value = False Then
    Exit Sub
End If
If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If

If OptT1.Value = True And mod1.DName = "乔继敏" And Lc = 5 Then
    XJF = MsgBox("现金流是否OK?", vbQuestion + vbYesNoCancel + vbDefaultButton3, "Hello")
    If XJF = vbCancel Then
        Exit Sub
    End If
End If
frmQm.Visible = False
        timZm = 3 '签字
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "合同执行"
        mod1.cmd.Parameters("@NBLX") = "付款审核"
        mod1.cmd.Parameters("@bh") = Zid
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = Uid
        mod1.cmd.Parameters("@mt2") = ""
        mod1.cmd.Parameters("@mt3") = lblXmmc.Caption
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

        mod1.cmd.Parameters("@mlt1") = txtQM.Text '评审建议
        mod1.cmd.Parameters("@mlt2") = ""
        mod1.cmd.Parameters("@mlt3") = ""
        mod1.cmd.Parameters("@mlt4") = ""
        mod1.cmd.Parameters("@mlt5") = ""
        If Lc = 0 Then Exit Sub
        mod1.cmd.Parameters("@mm1").Value = Lc
        If Fwid = 0 Then Exit Sub
        mod1.cmd.Parameters("@mm2").Value = Fwid
        If Amount = 0 Then Exit Sub
        mod1.cmd.Parameters("@mm3") = Amount
        mod1.cmd.Parameters("@mm4") = 0
        mod1.cmd.Parameters("@mm5") = 0
        mod1.cmd.Parameters("@mm6") = 0
        mod1.cmd.Parameters("@mm7") = 0
        mod1.cmd.Parameters("@mm8") = 0
        mod1.cmd.Parameters("@mm9") = 0
        If Fid = 0 Then Exit Sub
        mod1.cmd.Parameters("@mm10").Value = Fid
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
        mod1.cmd.Parameters("@mb2") = XJF
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

Private Sub cmdMod_Click()
If mod1.DName = "张军" Or mod1.DName = "徐瑛" Or mod1.DName = "" Then
    frmED.Visible = True
    cmdSave.Enabled = True
End If
End Sub

Private Sub cmdNQ_Click()

Dim tt As String
Dim oo As Integer

Dim ii As Integer


On Error Resume Next


dtgN.Row = 2: dtgN.Col = 1
If Val(dtgN.Text) = 0 Then Exit Sub



If LCRen <> mod1.DName Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If

If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If


    

    
    dtgN.Row = dtgN.Rows - 1
    dtgN.Col = 10
    Fid = Val(dtgN.Text)
    Lc = 1
    For oo = 2 To 8
        dtgN.Col = oo
        If dtgN.Text = "" Then
            Exit For
        End If
        Lc = Lc + 1
    Next
    dtgN.Col = 1
    Amount = Val(dtgN.Text)
    If Lc = 1 Then   '报销人只能签字，不能驳回。
        optT2.Enabled = False
        OptT1.Value = True
    Else
        optT2.Enabled = True
        OptT1.Value = False
        optT2.Value = False
    End If
    frmQm.Visible = True
    cmdDing.Enabled = True
    Exit Sub
End Sub

Private Sub cmdSave_Click()
Dim tt As String
Dim oo As Integer
On Error Resume Next
cmdSave.Enabled = False
timZm = 5 '保存
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "合同执行"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = Zid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mm1") = 0
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mm3") = 0
    mod1.cmd.Parameters("@mm4") = 0
    mod1.cmd.Parameters("@mm5") = 0
    mod1.cmd.Parameters("@mm6") = 0
    mod1.cmd.Parameters("@mm7") = 0
    mod1.cmd.Parameters("@mm8") = 0
    mod1.cmd.Parameters("@mm9") = 0
    mod1.cmd.Parameters("@mlt1") = Trim(txtBz.Text)
    mod1.cmd.Parameters("@mb1") = 0
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
        If timZm = 5 Then '保存
            cmdSave.Enabled = True
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
End Sub

Private Sub cmdZX_Click()
Dim tt As String
Dim YYY As Long
Dim hg As Single
Dim oo As Integer
On Error Resume Next

timZm = 2 '确认执行
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "合同执行"
    mod1.cmd.Parameters("@NBLX") = "确认执行"
    mod1.cmd.Parameters("@bh") = Zid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mm1") = Me.Fwid
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
        If timZm = 3 Then '保存

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

End Sub

Private Sub dtgMa_Click()
Dim oo As Integer
dtgN.Rows = dtgMa.Rows
dtgN.Row = dtgMa.Row
dtgN.Col = 9
If dtgN.Text = "True" Then
    cmdNQ.Visible = False
Else
    cmdNQ.Visible = True
End If
For oo = 2 To 7
    dtgN.Col = oo
    Lc = Lc + 1
    If dtgN.Text = "" Then
        Exit For
    End If
Next
frmQm.Visible = False
dtgN.Col = 10
Fid = Val(dtgN.Text)
dtgN.Col = 1
Amount = Val(dtgN.Text)
If Fid = 0 Then Exit Sub
Call QMBound(Fid)

End Sub

Private Sub dtgP_Click()
'frmQm.Visible = False
End Sub

Private Sub Form_Click()
frmQm.Visible = False
End Sub

Private Sub Form_Load()
Dim oo As Integer
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
dtgMa.Cols = 12
dtgN.Cols = 11
dtgMa.ColWidth(0) = -1
dtgMa.ColWidth(1) = 1200
dtgMa.Col = 0: dtgMa.CellFontBold = True
dtgMa.Col = 1: dtgMa.CellFontBold = True
For oo = 2 To 8
    dtgMa.ColWidth(oo) = 1300
    dtgMa.CellFontBold = True
Next
dtgMa.ColWidth(9) = 0
dtgMa.ColWidth(10) = 0
dtgMa.ColWidth(11) = 3480
dtgMa.Col = 11: dtgMa.CellFontBold = True
frmQm.Left = 6840
frmQm.Top = 7410
End Sub

Public Sub QMBound(Gid As Long)
Dim Ra: Dim La
Dim ii As Integer: Dim oo As Integer
Dim tt As String
On Error GoTo meERR2

tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & Gid & "' and yid=64 order by pid desc"
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
meERR2:
MsgBox "出错!"
End
End Sub
Public Sub Qing()
lblXmmc.Caption = ""
lblHtbh.Caption = ""
lblXz.Caption = ""
lblYwy.Caption = ""
dtgMa.Clear
txtBz.Text = ""
lblJe.Caption = ""
lblZtime.Caption = ""
frmED.Visible = False
Fwid = 0
LCUid = ""
LCRen = ""
Hid = 0
Pwf = False
Zid = 0
Uid = ""
ZT = ""
cmdSave.Enabled = False
cmdZx.Visible = False
End Sub

Public Sub dtgFF()

dtgMa.Row = 0
dtgMa.Col = 0: dtgMa.Text = "分笔": dtgMa.CellFontBold = True
dtgMa.Col = 1: dtgMa.Text = "金额": dtgMa.CellFontBold = True
dtgMa.Col = 2: dtgMa.Text = "付款申请人": dtgMa.CellFontBold = True
dtgMa.Col = 3: dtgMa.Text = "审核人1": dtgMa.CellFontBold = True
dtgMa.Col = 4: dtgMa.Text = "审核人2": dtgMa.CellFontBold = True
dtgMa.Col = 5: dtgMa.Text = "审核人3": dtgMa.CellFontBold = True
dtgMa.Col = 6: dtgMa.Text = "审核人4": dtgMa.CellFontBold = True
dtgMa.Col = 7: dtgMa.Text = "审核人5": dtgMa.CellFontBold = True
dtgMa.Col = 8: dtgMa.Text = "审核人6": dtgMa.CellFontBold = True
dtgMa.Col = 11: dtgMa.Text = "供应商": dtgMa.CellFontBold = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Cancel = True
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub

Public Sub Bound(Fwid As Long, Zid As Long)
Dim tt As String
Dim Ra, Rb, RC, RD
Dim Lc As Integer
If Fwid > 0 Then
    tt = "declare @hid int,@zid int,@lcUid nvarchar(10);" & _
        "select hid,uid,xz,ztime,pwf,pwf,bz,lcuid,fwid,zid,zt from htzx where fwid=" & Fwid & ";" & _
        "select @hid=hid,@zid=zid,@lcuid=lcuid from htzx where fwid=" & Fwid & ";" & _
        "select xywy,khmc,htbh from htping where hid=@hid;" & _
        "SELECT dbo.htzFk.amount, dbo.htzFk.pwf, dbo.htzFk.fid, dbo.htzFk.qm1, dbo.htzFk.qm2, dbo.htzFk.qm3, dbo.htzFk.qm4, dbo.htzFk.qm5, dbo.htzFk.qm6, " & _
         " dbo.htzFk.qm7 , SD30301_豪曼制冷.dbo.l_trader.Name FROM dbo.htzFk LEFT OUTER JOIN" & _
        " SD30301_豪曼制冷.dbo.l_trader ON dbo.htzFk.traderId = SD30301_豪曼制冷.dbo.l_trader.traderid where dbo.htzfk.zid=@zid and dbo.htzfk.amount>0 order by dbo.htzfk.fid;" & _
        "select username from worker where userid=@lcuid"

Else
    tt = "declare @hid int,@zid int,@lcUid nvarchar(10);" & _
        "select hid,uid,xz,ztime,pwf,pwf,bz,lcuid,fwid,zid,zt from htzx where Zid=" & Zid & ";" & _
        "select @hid=hid,@zid=zid from htzx where zid=" & Zid & ";" & _
        "select xywy,khmc,htbh from htping where hid=@hid;" & _
        "SELECT dbo.htzFk.amount, dbo.htzFk.pwf, dbo.htzFk.fid, dbo.htzFk.qm1, dbo.htzFk.qm2, dbo.htzFk.qm3, dbo.htzFk.qm4, dbo.htzFk.qm5, dbo.htzFk.qm6, " & _
         " dbo.htzFk.qm7 , SD30301_豪曼制冷.dbo.l_trader.Name FROM dbo.htzFk LEFT OUTER JOIN" & _
        " SD30301_豪曼制冷.dbo.l_trader ON dbo.htzFk.traderId = SD30301_豪曼制冷.dbo.l_trader.traderid where dbo.htzfk.zid=@zid and dbo.htzfk.amount>0 order by dbo.htzfk.fid;" & _
        "select username from worker where userid=@lcuid"

End If
Set mod1.HTP = CreateObject("adodb.recordset")
On Error GoTo htzERR
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
'Set mod1.HTP = mod1.HTP.NextRecordset
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RC = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
    RD = mod1.HTP.GetRows
End If
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next
Lc = UBound(RC, 2) + 1
On Error Resume Next
lblXmmc.Caption = Rb(1, 0)
lblHtbh.Caption = Rb(2, 0)
lblXz.Caption = Ra(2, 0)
lblYwy.Caption = Rb(0, 0)
txtBz.Text = Ra(6, 0)
lblZtime.Caption = Ra(3, 0)
Me.Fwid = Ra(8, 0)
LCUid = Ra(7, 0)
Hid = Ra(0, 0)
Pwf = Ra(4, 0)
Me.Zid = Ra(9, 0)
ZT = Ra(10, 0)
Uid = Ra(1, 0)
lblJe.Caption = 0
LCRen = RD(0, 0)
Call dtgPFF

If ZT = "未执行" Then
    cmdZx.Visible = True
Else
    cmdZx.Visible = False
End If
If Lc > 0 Then
    Call MaBound(RC, Lc)
End If
If Fid > 0 Then
Call QMBound(Fid)
End If
If mod1.DName = "张军" Then
    cmdZx.Visible = True
End If
Exit Sub
htzERR:
MsgBox "出错!"
frmZu.Enabled = True
'End

End Sub

Private Sub timQuit_Timer()
Dim oo As Integer
Dim ii As Integer
Dim tt As String
Dim Ra
Dim La As Integer
On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0

If timZm = 1 Or timZm = 6 Then '添加付款,删除付款
    tt = "SELECT dbo.htzFk.amount, dbo.htzFk.pwf, dbo.htzFk.fid, dbo.htzFk.qm1, dbo.htzFk.qm2, dbo.htzFk.qm3, dbo.htzFk.qm4, dbo.htzFk.qm5, dbo.htzFk.qm6, " & _
         " dbo.htzFk.qm7 , SD30301_豪曼制冷.dbo.l_trader.Name FROM dbo.htzFk LEFT OUTER JOIN" & _
        " SD30301_豪曼制冷.dbo.l_trader ON dbo.htzFk.traderId = SD30301_豪曼制冷.dbo.l_trader.traderid where dbo.htzfk.zid=" & Me.Zid & " and dbo.htzfk.amount>0 order by dbo.htzfk.fid"
    'tt = "select amount,pwf,fid,qm1,qm2,qm3,qm4,qm5,qm6,qm7 from htzFk where zid=" & Me.Zid & " order by fid"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2) + 1
    Call MaBound(Ra, La)
    frmED.Visible = False
ElseIf timZm = 2 Then '确认执行
    cmdZx.Visible = False
    If Dialog.Visible = True Then
        Call mod1.refEnvent(1)
    End If
ElseIf timZm = 3 Then '签字
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
    timQuit.Enabled = False

    If Dialog.Visible = True Then
        Call mod1.refEnvent(1)
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

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2,mt3,mt5 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 5 Then
        cmdSave.Enabled = False
    
    ElseIf timZm = 3 Then
        frmQm.Visible = False
        Lc = mod1.WP.Fields("mm1").Value
        Fwid = mod1.WP.Fields("mm2").Value
        LCRen = mod1.WP.Fields("mt1").Value
        LCUid = mod1.WP.Fields("mt2").Value
        LZw = mod1.WP.Fields("mt3").Value
            lblTX.Caption = "流程至" & LZw & ": " & LCRen
        dtgMa.Row = dtgMa.Rows - 1: dtgN.Row = dtgMa.Row
        If OptT1.Value = True Then
            dtgMa.Col = Lc: dtgN.Col = Lc
            dtgMa.Text = mod1.DName: dtgN.Text = mod1.DName
        Else
            For oo = 2 To 8
                dtgMa.Col = oo: dtgN.Col = oo
                dtgMa.Text = "": dtgN.Text = ""
            Next
        End If
       Call QMBound(Fid)
       If mod1.WP.Fields("mt5").Value = "1" Then
            dtgMa.Col = 9: dtgN.Col = 9
            dtgMa.Text = "True": dtgN.Text = "True"
            lblJe.Caption = Val(lblJe.Caption) + Amount
       End If
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    timWait.Enabled = False
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
'''''    If timZm = 2 Then
'''''        cmdSave.Enabled = False
'''''    ElseIf timZm = 11 Then
'''''        txtHtbh.Text = ""
'''''        lblHtxz.Caption = ""
'''''    End If
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0

    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub

Public Sub MaBound(Ra, La As Integer)
Dim oo As Integer
Dim Je As Single: Dim Lje As Single
dtgMa.Rows = La + 1
dtgN.Rows = dtgMa.Rows
dtgN.Cols = dtgMa.Cols

For oo = 1 To La
    dtgMa.Row = oo: dtgN.Row = oo
    dtgMa.Col = 0: dtgN.Col = 0
    dtgMa.Text = "第" & Str(oo) & "笔": dtgN.Text = "第" & Str(oo) & "笔"
    dtgMa.Col = 1: dtgN.Col = 1
    dtgMa.Text = Ra(0, oo - 1): dtgN.Text = dtgMa.Text 'amount
    
    dtgMa.Col = 9: dtgN.Col = 9
    dtgMa.Text = Ra(1, oo - 1): dtgN.Text = dtgMa.Text 'Pwf
    If dtgN.Text = "True" Then
        dtgN.Col = 1
        Lje = Val(dtgN.Text)
        Je = Je + Lje
    End If
    dtgMa.Col = 10: dtgN.Col = 10
    dtgMa.Text = Ra(2, oo - 1): dtgN.Text = dtgMa.Text 'fid
    dtgMa.Col = 11: dtgN.Col = 11
    dtgMa.Text = Ra(10, oo - 1): dtgN.Text = dtgMa.Text 'prader
    dtgMa.Col = 2: dtgN.Col = 2: dtgMa.Text = Ra(3, oo - 1): dtgN.Text = Ra(3, oo - 1)
    dtgMa.Col = 3: dtgN.Col = 3: dtgMa.Text = Ra(4, oo - 1): dtgN.Text = Ra(4, oo - 1)
    dtgMa.Col = 4: dtgN.Col = 4: dtgMa.Text = Ra(5, oo - 1): dtgN.Text = Ra(5, oo - 1)
    dtgMa.Col = 5: dtgN.Col = 5: dtgMa.Text = Ra(6, oo - 1): dtgN.Text = Ra(6, oo - 1)
    dtgMa.Col = 6: dtgN.Col = 6: dtgMa.Text = Ra(7, oo - 1): dtgN.Text = Ra(7, oo - 1)
    dtgMa.Col = 7: dtgN.Col = 7: dtgMa.Text = Ra(8, oo - 1): dtgN.Text = Ra(8, oo - 1)
    dtgMa.Col = 8: dtgN.Col = 8: dtgMa.Text = Ra(9, oo - 1): dtgN.Text = Ra(9, oo - 1)
Next
lblJe.Caption = Round(Je, 2)
dtgN.Row = dtgN.Rows - 1
dtgN.Col = 9
If dtgN.Text = "True" Then
    cmdNQ.Visible = False
Else
    cmdNQ.Visible = True
End If
dtgN.Col = 10
Fid = Val(dtgN.Text)
dtgN.Col = 1
Amount = Val(dtgN.Text)
End Sub

Private Sub txtBz_Click()
frmQm.Visible = False
End Sub


Private Sub txtTrader_DblClick()
frmTrader.Show
frmTrader.ZOrder 0
End Sub


