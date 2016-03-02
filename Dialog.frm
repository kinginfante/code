VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Dialog 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "共有N件事情等着您来处理"
   ClientHeight    =   3315
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   10050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdZZ 
      BackColor       =   &H00C0C0FF&
      Caption         =   "事务人"
      Height          =   315
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2880
      Width           =   915
   End
   Begin VB.CommandButton cmdR 
      Caption         =   "转单"
      Height          =   345
      Left            =   8640
      TabIndex        =   5
      Top             =   2430
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   285
      Left            =   8160
      TabIndex        =   15
      Top             =   3030
      Visible         =   0   'False
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   503
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "修正提交"
      Height          =   285
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "如果您觉得某项流程不应该跳至你处,可以将它提交至马晓聪处来进行修正."
      Top             =   2070
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdBJ 
      BackColor       =   &H00C0E0FF&
      Caption         =   "标记""非重要"""
      Height          =   315
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "将选中的任务级别设成一般,可以使你的事务列表中清楚显示您急于待办的事项"
      Top             =   1680
      Width           =   1275
   End
   Begin VB.OptionButton opt2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "一般"
      Height          =   210
      Left            =   9330
      TabIndex        =   12
      Top             =   720
      Width           =   675
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "重要"
      Height          =   195
      Left            =   8670
      TabIndex        =   11
      Top             =   720
      Value           =   -1  'True
      Width           =   675
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFC0C0&
      Caption         =   "查  询"
      Height          =   315
      Left            =   4530
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Width           =   1125
   End
   Begin VB.ComboBox comZ 
      Height          =   300
      Left            =   2430
      TabIndex        =   9
      Top             =   2880
      Width           =   2025
   End
   Begin VB.ComboBox comNR 
      Height          =   300
      ItemData        =   "Dialog.frx":0000
      Left            =   1080
      List            =   "Dialog.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdRef 
      BackColor       =   &H00FFC0C0&
      Caption         =   "刷新"
      Height          =   285
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   990
      Width           =   1275
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0FFFF&
      Caption         =   "打开"
      Height          =   285
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   1275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgDi 
      Height          =   2805
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   4948
      _Version        =   393216
      BackColor       =   12648384
      Cols            =   11
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0FFC0&
      Caption         =   "关闭"
      Height          =   285
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   1275
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除"
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblZZ 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7050
      TabIndex        =   16
      Top             =   2940
      Width           =   645
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查询分类:"
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   2910
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   30
      Left            =   30
      TabIndex        =   6
      Top             =   3000
      Width           =   30
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tlx As String '打开单子类型
Dim Tbh As Double '打开单子编号
Public OBF As Boolean '打开单子正确否
Public Fwid As Long
Public BHF As Boolean '是否为驳回单据
Option Explicit

Private Sub cmdBJ_Click()
Dim Odate As Date
Dim tt As String
Dim LX As String
Dim Ra
On Error Resume Next
On Error GoTo Dia1
dtgN.Row = dtgDi.Row
dtgN.Col = 0
LX = dtgN.Text
''''''''If LX = "合同执行通知" Then Exit Sub
If LX = "" Then Exit Sub
dtgN.Col = 2
Odate = dtgN.Text
'''''If DateDiff("d", Odate, mod1.DQda) > 13 Or LX = "成本追加通知" Or LX = "成本追加录速达" Or LX = "合同原件收集" Or LX = "财务到帐" Then
    dtgN.Col = 10
    Fwid = dtgN.Text
'''''    If Fwid = 0 Then Exit Sub
'''''    tt = "update NewFuwu set delf=0 where fwid=" & Fwid
'''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''    Call mod1.refEnvent(1)
'''''Else
'''''    MsgBox "过两周后,才能将此单列为非重要:)"
'''''End If
   '属于自己的事务列表，才能标记为非重要
If lblZZ.ToolTipText <> mod1.DHid Then Exit Sub
'if (mod1.DName = "" Or mod1.DName = "朱婷婷" Or mod1.DName = "陈文超") And (LX = "新合同执行通知" Or LX = "成本追加通知") Then Exit Sub
If mod1.DName = "陈文超" And (LX = "新合同执行通知" Or LX = "成本追加通知") Then Exit Sub

'''''''''    dtgN.Col = 10
'''''''''    Fwid = dtgN.Text
'''''''''    If Fwid = 0 Then Exit Sub
'''''''''
'''''''''    tt = "select uid from newfuwu where fwid=" & Fwid
'''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''''''    Ra = mod1.HTP.GetRows
'''''''''    mod1.HTP.Close
'''''''''    Set mod1.HTP = Nothing
'''''''''    If Ra(0, 0) <> mod1.DHid Then Exit Sub
    
    tt = "update NewFuwu set delf=0 where fwid=" & Fwid
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Call mod1.refEnvent(1)
Exit Sub
Dia1:
MsgBox "网络故障!"
End Sub

Private Sub cmdClose_Click()
Dialog.Visible = False
frmZu.Enabled = True
frmZu.TBa.Buttons(4).Value = tbrUnpressed
End Sub

Private Sub cmdDel_Click()
Dim tt As String
On Error Resume Next
dtgDi.Col = 10
Fwid = dtgDi.Text
If Fwid = 0 Then Exit Sub
If mod1.DName <> "马晓聪" Then Exit Sub
tt = "delete from NewFuwu where fwid=" & Fwid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
mod1.HTP.Close
Call mod1.refEnvent(1)

End Sub

Private Sub cmdOpen_Click()
Dim oo As Integer
'MsgBox "正在建设中,双击列表中的记录可以替代此功能!"
Dim Tlx As String
Dim LX As Integer '旧版本的打开类型
Dim tt As String
Dim Tbh As Long
Dim ZL As String '询价单的种类
Dim XLX As Boolean
Dim Htbh As String '合同评审单的编号
Dim Ny As Single
Dim Pwf As Boolean
Dim Lei As String
Dim Ra: Dim Uid As String: Dim FR As Date
Dim NewF As Integer
Dim htRow As Integer
On Error Resume Next
dtgN.Row = dtgDi.Row
BHF = False
dtgN.Col = 0
Tlx = dtgN.Text
dtgN.Col = 6
Tbh = dtgN.Text
dtgN.Col = 0
Tlx = dtgN.Text
dtgN.Col = 1
Lei = dtgN.Text
dtgN.Col = 5
If dtgN.Text = "驳回单据" Then
    BHF = True
End If
If Tlx = "" Then Exit Sub
'MsgBox TLx
'类型转换,以适合旧版本的唯一打开者子程序.
Select Case Tlx
    Case "合同评审单"
        LX = 1
    Case "报销单"
        LX = 2
        mod1.BTZ = 23
    Case "配料单"
        LX = 5
    Case "项目资料"
        LX = 6
End Select
dtgN.Col = 6
Tbh = dtgN.Text
'MsgBox TBh
Dialog.Enabled = False

If Tlx = "确定工程组长" Then
    Tlx = "询价单"
End If

Select Case Tlx
    Case "新项目资料"
    Dim Kid As Long
    Dim xid As Long
    'dtgKH.Col = 2
    xid = Tbh
    

    wbDN.Visible = False
    Me.MousePointer = 11
    mod1.BTZ = 1
    Call mod1.xmQing
    Call mod1.khQing
    Call mod1.xmBound(xid)
    wbDN.lblKid.Caption = wbDN.lblYz.Tag
    Call mod1.khBound(wbDN.lblYz.Tag, "yz")

    wbDN.frmJE.Visible = False

    wbDN.Left = 0
    wbDN.Top = 0
    wbDN.cmdMod.Enabled = False
    wbDN.cmdSave.Enabled = False
    Me.MousePointer = 0
    wbDN.tabKh.Tab = 0

    wbDN.tabKh.TabEnabled(2) = True
    wbDN.tabKh.TabEnabled(0) = True
    

    

    wbDN.modFi = False

    Me.MousePointer = 0
    wbDN.cmdSave.Enabled = False
    wbDN.tabKh.Enabled = True

    wbDN.khAdd = False
    '打开项目后,默认的打开客户为项目资料
    wbDN.optYz.Value = True
    wbDN.frmGL.Visible = False
    frmWait.Visible = False
    wbDN.Visible = True
    wbDN.cmdQing.Enabled = False
    wbDN.cmdNew.Enabled = False
    wbDN.cmdRadd.Enabled = False
    wbDN.cmdRdel.Enabled = False
    If wbDN.comXyxz.Text = "物业公司" Then
        wbDN.frmGL.Visible = True
    End If
    
    Case "新货品资料"
        tt = "select pid from nlpmxc where bh='" & dtgN.Text & "'"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        
        Call frmHPZL.Qing
        Tbh = Ra(0, 0)
'''        Call frmHPZL.BoundL1
'''        Call frmHPZL.dtgL2FF
'''        Call frmHPZL.dtgL3FF
      Call frmHPZL.Bound(Tbh)
        frmHPZL.Show
        frmHPZL.ZOrder 0
    Case "供应商资料"
        Call frmGyDetail.Qing
        Call frmGyDetail.Bound(Tbh)
        frmGyDetail.cmdSave.Enabled = False
        frmGyDetail.Show
        frmGyDetail.ZOrder 0
    Case "新合同执行通知"
            Call FmxcNew.Bound(Tbh)
            FmxcNew.Show
            FmxcNew.ZOrder 0
            If mod1.DName <> "乔继敏" And mod1.DName <> "于晓静" And mod1.DName <> "王绣霞" And mod1.DName <> "朱婷婷" Then
            Call FmxcNew.Xian
            End If
            Exit Sub
    Case "录订单通知"
            mod1.BTZ = 6
        dtgN.Col = 6
        tt = "select newF from htping where hid=" & Tbh
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        If Ra(0, 0) = 6 Or Ra(0, 0) = 8 Then
            Call FmxcNew.Bound(Tbh)
            FmxcNew.Show
            FmxcNew.ZOrder 0
            Exit Sub
        End If
            Call modNewHT.NewMQing
            Call modNewHT.NewB(Tbh)
            FMXC.lblMQM(0).Visible = True
            FMXC.lblMTm(0).Visible = True
            FMXC.cmdMQm(0).Visible = True
    Case "合同原件收集"
            tt = "select newF from htping where hid=" & Tbh
            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
            Ra = mod1.HTP.GetRows
            mod1.HTP.Close
            Set mod1.HTP = Nothing
            If Ra(0, 0) = 6 Or Ra(0, 0) = 8 Then
                Call FmxcNew.Bound(Tbh)
                FmxcNew.Show
                FmxcNew.ZOrder 0
                Exit Sub
            End If
            Call modNewHT.NewMQing
            Call modNewHT.NewB(Tbh)
            FMXC.lblMQM(0).Visible = True
            FMXC.lblMTm(0).Visible = True
            FMXC.cmdMQm(0).Visible = True
    Case "财务到帐"
        Call fmxcY.Bound(Val(Tbh))
        fmxcY.Show
        fmxcY.ZOrder 0
    Case "成本追加录速达"

        Call fmxcZJ.Bound(Val(Tbh))
        fmxcZJ.Show
        fmxcZJ.ZOrder 0
    Case "成本追加通知"

        Call fmxcZJ.Bound(Val(Tbh))
        fmxcZJ.Show
        fmxcZJ.ZOrder 0
    Case "成本追加单"

        Call fmxcZJ.Bound(Val(Tbh))
        fmxcZJ.Show
        fmxcZJ.ZOrder 0
    Case "合同执行通知"
    dtgN.Col = 10
    Tbh = Val(dtgN.Text)
        Call frmHtz1.Qing
        Call frmHtz1.Bound(Tbh, 0)
        Call frmHtz1.dtgFF
        
        frmHtz1.Show
    Case "付款审核"
        dtgN.Col = 10
        Tbh = Val(dtgN.Text)
        Call frmHtz1.Qing
        Call frmHtz1.Bound(Tbh, 0)
        Call frmHtz1.dtgFF
        
        frmHtz1.Show
    Case "工作报告"
    mod1.BTZ = 4
    tt = "select uid,fr from SalesReport where gid=" & Val(Tbh)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Uid = Ra(0, 0): FR = Ra(1, 0)
    'Call frmGZbN1.Qing
    'Call frmGZbN1.Bound(Uid, FR)
    'frmGZbN1.Show
    Case "报销单"
        mod1.BTZ = 23
        If mod1.DKZ(Tbh, LX) = True Then
                MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
                Dialog.Enabled = True
                Exit Sub
        End If
        
        frmFYBX.Show
        Call ModBx.FyQing
        Call ModBx.fydBound(Val(Tbh))
'''''        frmFYBX.lblLcRen.Caption = mod1.DName
'''''        frmFYBX.lblLcUid.Caption = mod1.DHid
        If BHF = True Then

            'Pje.Show
            tt = "select bz from pizu where bh='" & Tbh & "' and yid=" & frmFYBX.lblNlb.Caption & " order by trq desc"
            Set Pje.adoPje = CreateObject("adodb.recordset")
            Pje.adoPje.Close
            Pje.adoPje.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'            Set Pje.dtgPje.DataSource = Pje.adoPje
'            Pje.txtXQ.Text = ""
            frmFYBX.lblTX.Caption = "驳回原因:" & Pje.adoPje.Fields("bz").Value & "，详情请参见评审建议！"
            frmFYBX.lblTX.Visible = True
        End If
    Case "奖金"

    
        Dim QFF As Boolean
        mod1.BTZ = 23
        
        frmYjBx.Visible = False
        Call frmYjBx.yjBXQing
        Call frmYjBx.Bound(Val(Tbh))
        frmYjBx.Show
        Exit Sub
YZERR1:
        MsgBox "网络故障，请再试一次，或者重启程序！"
        Exit Sub
    Case "项目资料"
        mod1.BTZ = 1
        tt = "Select xid,kid from khren where rid=" & Val(Tbh)
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        
'''''        If mod1.DKZ(mod1.HTP.Fields("xid").Value, 6) = True Then
'''''        MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
'''''        Dialog.Enabled = True
'''''        Exit Sub
'''''        End If

          wbDN.Visible = False
          Me.MousePointer = 11
'''''          '记录打开日志
'''''          Call mod1.zhuDa(3, mod1.HTP.Fields("xid").Value)
          Call mod1.xmQing
          Call mod1.khQing
          
          Call mod1.khFuBound(mod1.HTP.Fields("kid").Value, mod1.HTP.Fields("xid").Value, Tbh)
        
          wbDN.cmdMod.Enabled = False
          wbDN.cmdSave.Enabled = False
          wbDN.tabKh.Tab = 1
'          wbDN.cmdRadd.Enabled = False
'          wbDN.cmdNew.Enabled = False
          wbDN.khAdd = False
          frmWait.Visible = False
          wbDN.Visible = True
          'wbDN.adoRen.Recordset.Move 0
          Me.MousePointer = 0
          If wbDN.lblYwy.Caption = mod1.DName Or wbDN.lblXywy.Caption = mod1.DName Then
              wbDN.cmdMod.Enabled = True
          Else
              wbDN.cmdMod.Enabled = False
          End If
          wbDN.lblLcRen.Caption = mod1.DName
          wbDN.lblLcUid.Caption = mod1.DHid
          wbDN.cmdMod.Enabled = True
     Case "询价单"
        mod1.BTZ = 36
        Me.Enabled = False
        frmWait.Visible = True
        frmWait.ZOrder 0
        frmWait.Refresh
        'If mod1.DName = "谢雪梅" Or mod1.MName = "马晓聪" Then

        frmWBXJ.Visible = False
        tt = "select Zl,LX,htrow from xunjiaD where bid=" & Tbh
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        ZL = mod1.HTP.Fields("zl").Value
        XLX = mod1.HTP.Fields("lx").Value
        htRow = mod1.HTP.Fields("htrow").Value
        If htRow > 0 Or ZL = "询价指令" Then
            Call FmxcXJ.Bound(Tbh)
            FmxcXJ.Show
            FmxcXJ.ZOrder 0
            Exit Sub
        End If
        If ZL = "人工" Or ZL = "维保" Or ZL = "大修" Or ZL = "其他人工" Or ZL = "压缩机维修保养" Or ZL = "中介业务" Or ZL = "分包" Then
            Call frmWBXX.Qing
            Call frmWBXX.Bound(Tbh)
            frmWBXX.Show
            frmWBXX.ZOrder 0
            Exit Sub
        End If
        If Val(Tbh) > 8113 And ZL = "维保" Then
            Call frmWBXNew.Qing
            Call frmWBXNew.Bound(Tbh)
            frmWBXNew.Show
            frmWBXNew.ZOrder 0
            Dialog.dtgN.Col = 10
            frmWBXNew.lblFwid = Val(dtgN.Text)
            Exit Sub
        End If
        Call modBJD.BJDWBQing
        Call modBJD.BJDGXQing

        If (ZL = "维保" Or ZL = "大修" Or ZL = "工程分包" Or ZL = "水处理") And XLX = True Then

                    Call frmWBXNew.Qing
                    Call frmWBXNew.Bound(Val(Tbh))
                    frmWBXNew.frmM1.Visible = False
                    frmWBXNew.Show
                    frmWBXNew.ZOrder 0
        ElseIf (ZL = "维保" Or ZL = "大修" Or ZL = "工程分包" Or ZL = "水处理") And XLX = False Then '直接打开维保大修询价单的购销询价表
            Call modBJD.BJDGXQing
            Call modBJD.BJDGDBound(Val(Tbh))
            Call modBJD.gxbjLocked
            tt = "select bid from xunjiaOld where oid=" & Val(frmGXBj.lblOid.Caption) & " order by bid"
            frmGXBj.adoOid.Close
            frmGXBj.adoOid.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            frmGXBj.adoOid.MoveLast
            If frmGXBj.adoOid.RecordCount > 1 Then
                frmGXBj.cmdRight.Enabled = False
                frmGXBj.cmdLeft.Enabled = True
            Else
                frmGXBj.cmdRight.Enabled = False
                frmGXBj.cmdLeft.Enabled = False
            End If
        
            frmWait.Visible = False
            frmGXBj.Visible = True
            frmGXBj.ZOrder 0
            frmGXBj.cmdMod.Enabled = True
            frmGXBj.cmdSave.Enabled = False
''''            frmGXBj.lblLcRen.Caption = mod1.DName
''''            frmGXBj.lblLcUid.Caption = mod1.DHid
        Else
            If mod1.Mname = "马晓聪" Or mod1.DName = "谢雪梅" Or mod1.DName = "货品录入员" Or mod1.DName = "杨晓刚" Then
                Call frmGxbjNew.Initialize
                Call frmGxbjNew.Bound(Val(Tbh))
                mod1.BTZ = 36
                frmWait.Visible = False
                frmGxbjNew.Visible = True
                frmGxbjNew.ZOrder 0
                frmGxbjNew.cmdMod.Enabled = True
                frmGxbjNew.cmdSave.Enabled = False
                Exit Sub
            End If
            Call modBJD.BJDGXQing
            Call modBJD.BJDBound(Val(Tbh), ZL)
            Call frmGXBj.dtgMaFF
            Call modBJD.gxbjLocked
            If frmGXBj.lblYwy = "谢雪梅" Or Tbh > 10058 Then
                'frmGXBj.frmSD.Visible = True
                frmGXBj.frmCg.Top = 4740
                frmGXBj.dtgNew.Visible = True
                'frmGXBj.cmdPje.Visible = False
                frmGXBj.dtgP.Visible = True
            Else
                'frmGXBj.frmSD.Visible = False
                frmGXBj.frmCg.Top = 7620
                frmGXBj.dtgNew.Visible = False
                'frmGXBj.cmdPje.Visible = True
                frmGXBj.dtgP.Visible = False
            End If
''''''''''            tt = "select bid from xunjiaOld where oid=" & Val(frmGXBj.lblOid.Caption) & " order by bid"
''''''''''            frmGXBj.adoOid.Close
''''''''''            frmGXBj.adoOid.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''''''            frmGXBj.adoOid.MoveLast
''''''''''            If frmGXBj.adoOid.RecordCount > 1 Then
''''''''''                frmGXBj.cmdRight.Enabled = False
''''''''''                frmGXBj.cmdLeft.Enabled = True
''''''''''            Else
''''''''''                frmGXBj.cmdRight.Enabled = False
''''''''''                frmGXBj.cmdLeft.Enabled = False
''''''''''            End If
        
            frmWait.Visible = False
            frmGXBj.Visible = True
            frmGXBj.ZOrder 0
            frmGXBj.cmdMod.Enabled = True
            frmGXBj.cmdSave.Enabled = False
'''''            frmGXBj.lblLcRen.Caption = mod1.DName
'''''            frmGXBj.lblLcUid.Caption = mod1.DHid
''''''''''            If BHF = True Then
''''''''''
''''''''''                'Pje.Show
''''''''''                tt = "select bz from pizu where bh='" & TBh & "' and yid=" & frmFYBX.lblNlb.Caption & " order by trq desc"
''''''''''                Set Pje.adoPje = CreateObject("adodb.recordset")
''''''''''                Pje.adoPje.Close
''''''''''                Pje.adoPje.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
''''''''''    '            Set Pje.dtgPje.DataSource = Pje.adoPje
''''''''''    '            Pje.txtXQ.Text = ""
''''''''''                frmGXBj.lblTX.Caption = "驳回原因:" & Pje.adoPje.Fields("bz").Value & "，详情请参见评审建议！"
''''''''''                frmGXBj.lblTX.Visible = True
''''''''''            End If
        End If
     Case "确定工程组长"
                mod1.BTZ = 36
                Me.Enabled = False
                frmWait.Visible = True
                frmWait.ZOrder 0
                frmWait.Refresh

            Call frmWBXNew.Qing
            Call frmWBXNew.Bound(Tbh)
            Dialog.dtgN.Col = 11
            frmWBXNew.lblFwid = Val(dtgN.Text)
            frmWBXNew.Show
            frmWBXNew.ZOrder 0
            Exit Sub

     Case "报价单"
        mod1.BTZ = 37
        Me.Enabled = False
        frmWait.Visible = True
        frmWait.ZOrder 0
        frmWait.Refresh
        frmWbxjB.Visible = False
        tt = "select lx from baojiaD where baoid=" & Tbh
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        If mod1.HTP.Fields("lx").Value = True Then
            frmGxBiao.Visible = False
            Call modBJD.BaoJDWBQing
            Call modBJD.BaoJDBound(CInt(Tbh), mod1.HTP.Fields("LX").Value)
            
'            tt = "select * from baojiaOld where old=" & Val(frmWbxjB.lblOid.Caption) & " order by baoid"
'            frmWbxjB.adoOid.Close
'            frmWbxjB.adoOid.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'            If frmWbxjB.adoOid.RecordCount > 1 Then
'                frmWbxjB.cmdLeft.Enabled = True
'            End If
            frmWbxjB.adoOid.MoveLast
            frmWait.Visible = False
            frmWbxjB.Visible = True
            frmWbxjB.ZOrder 0
            'Dialog.Enabled = True
            frmWbxjB.lblLcRen.Caption = mod1.DName
            frmWbxjB.lblLcUid.Caption = mod1.DHid
            frmWbxjB.cmdMod.Enabled = True
        Else
            frmGxBiao.Visible = False
            Call modBJD.BaoJDGXQing
            Call modBJD.BaoJDBound(CInt(Tbh), mod1.HTP.Fields("LX").Value)

            tt = "select * from baojiaOld where old=" & Val(frmGxbjB.lblOid.Caption) & " order by baoid"
            frmGxbjB.adoOid.Close
            frmGxbjB.adoOid.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            If frmGxbjB.adoOid.RecordCount > 1 Then
                frmGxbjB.cmdLeft.Enabled = True
            End If
            frmGxbjB.adoOid.MoveLast
            frmWait.Visible = False
            frmGxbjB.Visible = True
            frmGxbjB.ZOrder 0
            'Dialog.Enabled = True
            frmGxbjB.lblLcRen.Caption = mod1.DName
            frmGxbjB.lblLcUid.Caption = mod1.DHid
            frmGxbjB.cmdMod.Enabled = True
        End If
     Case "销售日记"
        mod1.BTZ = 4
        Me.Enabled = False
        frmWait.Visible = True
        frmWait.ZOrder 0
        frmWait.Refresh
        frmGzNr.Visible = False
        tt = "select lb from xmgz where gid=" & Tbh
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        If mod1.HTP.Fields("lb").Value = True Then
            Call modXmGz.xmQing
            Call modXmGz.xmBound(CInt(Tbh))
            frmWait.Visible = False
            frmGzNr.Visible = True
            frmGzNr.ZOrder 0
            'Dialog.Enabled = True
        End If
        frmGzNr.lblLcRen.Caption = mod1.DName
        frmGzNr.lblLcUid.Caption = mod1.DHid
    Case "合同评审单"
        mod1.BTZ = 6
        dtgN.Col = 6
        tt = "select newF from htping where hid=" & Tbh
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        If Ra(0, 0) = 6 Or Ra(0, 0) = 8 Then
            Call FmxcNew.Bound(Tbh)
            FmxcNew.Show
            FmxcNew.ZOrder 0
            Exit Sub
        End If
        If Left(dtgN.Text, 2) = "HM" Then
            Htbh = dtgN.Text
            tt = "select hid,newf from htView where 合同编号='" & Htbh & "'"
            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
            Tbh = mod1.HTP.Fields("hid").Value
            NewF = mod1.HTP.Fields("newf").Value
            If mod1.DKZ(Tbh, LX) = True Then
                    MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
                    Dialog.Enabled = True
                    Exit Sub
            End If
            Me.Enabled = False
            frmWait.Visible = True
            frmWait.ZOrder 0
            frmWait.Refresh
            
            frmWbNew.Visible = False
            
            Call modHt.NewQing
    
            Call modHt.NewBound(Tbh)
            
            frmWbNew.Visible = True
            frmWbNew.lblLcRen.Caption = mod1.DName
            frmWbNew.lblLcUid.Caption = mod1.DHid
        ElseIf Val(Tbh) < 19345 Then
''''            FMXC.Show
''''            Exit Sub
            Call modNewHT.NewMQing
        
        
            Call modNewHT.NewMBound(Tbh)
            FMXC.lblMQM(0).Visible = True
            FMXC.lblMTm(0).Visible = True
            FMXC.cmdMQm(0).Visible = True
'            FMXC.lblLcRen.Caption = mod1.DName
'            FMXC.lblLcUid.Caption = mod1.DHid
        Else
            Call modNewHT.NewMQing
            Call modNewHT.NewB(Tbh)
            FMXC.lblMQM(0).Visible = True
            FMXC.lblMTm(0).Visible = True
            FMXC.cmdMQm(0).Visible = True
        End If
     Case "配料单"
        Dim Pmid As Long
        Dim POid As Long
        'mod1.BTZ = 4
        Me.Enabled = False
        frmWait.Visible = True
        frmWait.ZOrder 0
        frmWait.Refresh
        'frmPld.Visible = False
        dtgN.Col = 6
        
        Pmid = dtgN.Text
        If mod1.DKZ(Pmid, LX) = True Then
                MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
                Exit Sub
        End If
        
        Call modPld.PLDQing
        Call modPld.PLDBound(Pmid)
        
        tt = "select guid from pldMain where pmid=" & Pmid
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        POid = mod1.HTP.Fields("guid").Value
        
        '打开旧单子
        Set mod1.PldO = CreateObject("adodb.recordset")
        tt = "PldOldCount(" & POid & ")"
        mod1.PldO.Close
        mod1.PldO.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        
        If mod1.PldO.RecordCount > 0 Then
            mod1.PldO.MoveLast
            Call modPld.PldOldBound(mod1.PldO.Fields("Pmid").Value)
        
            'frmPld.cmdRight.Enabled = False
            'frmPld.cmdLeft.Enabled = True
            'frmPld.Height = 9750
        Else
            'frmPld.Height = 5895
        End If
'        frmPld.lblZT.Visible = True
'        frmPld.Visible = True
'        frmPld.ZOrder 0
        frmWait.Visible = False
        'frmPld.lblLcRen.Caption = mod1.DName
        'frmPld.lblLcUid.Caption = mod1.DHid

    Case "胡萝卜"
        Call HLB.HLBQing
        Call HLB.HLBBound(Tbh)
        HLB.Show

    Case "设置权限"
        Call frmRen.RenQing
        Call frmRen.RenBound(dtgN.Text)
        
        tt = "select userid as 工号,username as 姓名,qy as 区域,bm as 部门,userzw as 职务,nx as 工作年限 from worker where userid='" & dtgN.Text & "'"
        frmRen.adoRen.Close
        frmRen.adoRen.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Set frmRen.dtgRen.DataSource = frmRen.adoRen
        frmRen.Visible = True
        frmRen.ZOrder 0
'''''    Case "绩效考核"
'''''        Call b1.KPIQing
'''''        Call b1.KPIBound1(TBh, Lei)
        
End Select
End Sub











Private Sub cmdR_Click()
dtgDi.Col = 10
Fwid = dtgDi.Text
If Fwid = 0 Then Exit Sub
Set Ren.XForm = New Dialog

'Call mod1.RenXz("Dialog", Me, 0)

End Sub

Private Sub cmdRef_Click()
If opt1.Value = True Then
    Call mod1.refEnvent(1)
    cmdBJ.Enabled = True
Else
    Call mod1.refEnvent(0)
    cmdBJ.Enabled = False
End If
End Sub

Private Sub cmdSearch_Click()
Dim tt As String
If comNR.Text = "内容" Then

    mod1.ETT = "select LX as 事务类型,nr as 内容,RQ as 生成时间,ywy,uid,Lab as 身份职能,Bh as 编号,DxRen as 对象姓名,cf,crq,fwid from NewFu where (uid='" & lblZZ.ToolTipText & "' and not((lx='合同评审单' and lab='执行完毕确认') or (lx='配料单' and lab='成本结算'))) and nr like '%" & comZ.Text & "%'"

ElseIf comNR.Text = "类型" Then
    mod1.ETT = "select LX as 事务类型,nr as 内容,RQ as 生成时间,ywy,uid,Lab as 身份职能,Bh as 编号,DxRen as 对象姓名,cf,crq,fwid from NewFu where (uid='" & lblZZ.ToolTipText & "' and not((lx='合同评审单' and lab='执行完毕确认') or (lx='配料单' and lab='成本结算'))) and lx='" & comZ.Text & "'"
ElseIf comNR.Text = "编号" Then

    mod1.ETT = "select LX as 事务类型,nr as 内容,RQ as 生成时间,ywy,uid,Lab as 身份职能,Bh as 编号,DxRen as 对象姓名,cf,crq,fwid from NewFu where (uid='" & lblZZ.ToolTipText & "' and not((lx='合同评审单' and lab='执行完毕确认') or (lx='配料单' and lab='成本结算'))) and bh='" & comZ.Text & "'"
End If

        Dim RL
        Dim ul
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open (mod1.ETT & " order by rq desc"), mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        If mod1.HTP.BOF = False Then
            RL = mod1.HTP.GetRows
            ul = UBound(RL, 2)
        End If
        mod1.HTP.Close
        Set mod1.HTP = Nothing

        Call mod1.refEnt2(RL, ul)

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdZZ_Click()
If mod1.BmJl = False And mod1.Bm = "配送中心" And mod1.Mname <> "马晓聪" And mod1.DName <> "马晓聪" Then Exit Sub

Set Ren.XForm = New frmRen
Call mod1.RenXz("Dialog", Me, 0)
End Sub

Private Sub comNR_Click()
Dim oo As Integer
Dim tt As String
On Error Resume Next
If Me.Visible = False Then Exit Sub
If comNR.Text = "内容" Then
    comZ.Text = ""
    For oo = 15 To 0 Step -1
        comZ.RemoveItem oo
    Next
ElseIf comNR.Text = "类型" Then
    For oo = 15 To 0 Step -1
        comZ.RemoveItem oo
    Next
    comZ.AddItem "合同评审单"
    comZ.AddItem "报销单"
    comZ.AddItem "销售日记"
    comZ.AddItem "项目资料"
    comZ.AddItem "施工进度表"
    comZ.AddItem "询价单"
    comZ.AddItem "报价单"
End If
End Sub

Private Sub dtgDi_DblClick()
On Error Resume Next

Static Px As Boolean

'If dtgDi.Row = 1 Then
    If Px = True Then
        dtgDi.Sort = 2
        Px = False
    Else
        dtgDi.Sort = 1
        Px = True
    End If
'Else
'    MsgBox MGa.ColData(1)
'End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode
On Error Resume Next
Dim zz As Long
Dim L As Integer
Fwid = 0
If Shift = 6 And KeyCode = 82 Then
    If cmdDel.Visible = True Then
        cmdDel.Visible = False
        cmdR.Visible = False
    Else
        cmdDel.Visible = True
        cmdR.Visible = True
    End If
End If

'''''If KeyCode = 74 Then
''''''dtgDi.Col = 1
''''''    ZZ = 0
''''''    dtgDi.Row = 1
''''''    L = dtgDi.Row
''''''    ZZ = Val(dtgDi.Text)
''''''    Do While Not L > dtgDi.Rows + 2000
''''''            dtgDi.Row = dtgDi.Row + 1
''''''                L = L + 1
''''''        ZZ = ZZ + Val(dtgDi.Text)
''''''
''''''
''''''    Loop
'''''    ZZ = 0
'''''    AdoDi.MoveFirst
'''''    Do While Not AdoDi.EOF
'''''        ZZ = ZZ + AdoDi.Fields("内容").Value
'''''        AdoDi.MoveNext
'''''    Loop
'''''    MsgBox ZZ
'''''End If
End Sub

Private Sub Form_Load()
Dialog.Height = 3795
Dialog.Width = 10140
dtgDi.Row = 0
dtgDi.Col = 0
dtgDi.Text = "事务类型"
dtgDi.Col = 1
dtgDi.Text = "内容"
dtgDi.Col = 2
dtgDi.Text = "生成时间"
dtgDi.Col = 3
dtgDi.Text = "ywy"
dtgDi.Col = 4
dtgDi.Text = "uid"
dtgDi.Col = 5
dtgDi.Text = "身份职能"
dtgDi.Col = 6
dtgDi.Text = "编号"
dtgDi.Col = 7
dtgDi.Text = "对象姓名"
dtgDi.Col = 8
dtgDi.Text = "cf"
dtgDi.Col = 9
dtgDi.Text = "crq"
dtgDi.Col = 10
dtgDi.Text = "fwid"

dtgDi.ColWidth(0) = 1300
dtgDi.ColWidth(1) = 1500
dtgDi.ColWidth(2) = 2000
dtgDi.ColWidth(3) = 0
dtgDi.ColWidth(4) = 0
dtgDi.ColWidth(5) = 1000
dtgDi.ColWidth(6) = 1500
dtgDi.ColWidth(7) = 1000
dtgDi.ColWidth(8) = 0
dtgDi.ColWidth(9) = 0
dtgDi.ColWidth(10) = 0
cmdDel.Visible = False


End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
    Dialog.Visible = False
    frmZu.Enabled = True
    Cancel = True
    frmZu.TBa.Buttons(4).Value = tbrUnpressed
End If
End Sub

Private Sub OKButton_Click()

End Sub


Public Sub dtgDiFF()
Dim oo As Integer
Dim ii As Integer
On Error Resume Next

Dialog.dtgDi.Clear: Dialog.dtgN.Clear
Dialog.dtgDi.Row = 0
Dialog.dtgDi.Col = 0
Dialog.dtgDi.Text = "事务类型"
Dialog.dtgDi.Col = 1
Dialog.dtgDi.Text = "内容"
Dialog.dtgDi.Col = 2
Dialog.dtgDi.Text = "生成时间"
Dialog.dtgDi.Col = 3
Dialog.dtgDi.Text = "ywy"
Dialog.dtgDi.Col = 4
Dialog.dtgDi.Text = "uid"
Dialog.dtgDi.Col = 5
Dialog.dtgDi.Text = "身份职能"
Dialog.dtgDi.Col = 6
Dialog.dtgDi.Text = "编号"
Dialog.dtgDi.Col = 7
Dialog.dtgDi.Text = "对象姓名"
Dialog.dtgDi.Col = 8
Dialog.dtgDi.Text = "cf"
Dialog.dtgDi.Col = 9
Dialog.dtgDi.Text = "crq"
Dialog.dtgDi.Col = 10
Dialog.dtgDi.Text = "fwid"

End Sub
