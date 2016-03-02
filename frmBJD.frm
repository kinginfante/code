VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmCWBB 
   Caption         =   "财务报表"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15150
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   15150
   Begin VB.CommandButton cmdVnew 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查  询"
      Height          =   315
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8670
      Width           =   1935
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9270
      Top             =   8490
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10080
      Top             =   8490
   End
   Begin VB.CommandButton cmdJZ 
      BackColor       =   &H000000FF&
      Caption         =   "结 帐"
      Height          =   345
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8670
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgCol 
      Height          =   3645
      Left            =   6900
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   6429
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdCopy 
      BackColor       =   &H00C0FFFF&
      Caption         =   "复制 -> EXCEL"
      Height          =   285
      Left            =   3570
      TabIndex        =   5
      Top             =   8340
      Width           =   1935
   End
   Begin VB.CommandButton cmdV 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查  询"
      Height          =   315
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8700
      Width           =   1935
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "frmBJD.frx":0000
      Left            =   1050
      List            =   "frmBJD.frx":0013
      TabIndex        =   3
      Text            =   "团队费用"
      Top             =   8700
      Width           =   2385
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "导航"
      Height          =   585
      Left            =   14310
      Picture         =   "frmBJD.frx":0052
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8430
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBB 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   14631
      _Version        =   393216
      Rows            =   40
      Cols            =   15
      FixedRows       =   0
      FixedCols       =   0
      ForeColorSel    =   65535
      BackColorBkg    =   -2147483627
      BackColorUnpopulated=   65535
      GridColorUnpopulated=   8421376
      FillStyle       =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   15
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   315
      Left            =   540
      TabIndex        =   6
      Top             =   8310
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   8454016
      CustomFormat    =   "yyyy年-MM月"
      Format          =   55705603
      CurrentDate     =   38797
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   315
      Left            =   2010
      TabIndex        =   7
      Top             =   8310
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   8454016
      CustomFormat    =   "yyyy年-MM月"
      Format          =   55705603
      CurrentDate     =   38797
   End
   Begin VB.Label lblCell 
      BackColor       =   &H0000FFFF&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7110
      TabIndex        =   9
      Top             =   8640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "日期:"
      Height          =   225
      Left            =   30
      TabIndex        =   8
      Top             =   8370
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "报表类型"
      Height          =   255
      Left            =   30
      TabIndex        =   2
      Top             =   8730
      Width           =   975
   End
End
Attribute VB_Name = "frmCWBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoBB As ADODB.Recordset
Dim Orow As Long
Dim OCol As Long
Dim Crow As Long
Dim Ccol As Long
Dim Oc As Long
Dim timZm As Integer
Private Sub cmdAll_Click()
Dim tt As String
On Error Resume Next
If mod1.KhK = 1 Then
    tt = "select * from bjdV where bm='" & mod1.BM & "' order by 报价日期 desc"
ElseIf mod1.KhK = 2 Then
    If mod1.comId <> 0 Then
    Else
        tt = "select * from bjdV where comid=" & comId & " and not(bm='维销部3' or bm='产品部1' or bm='产品部2')  order by 报价日期 desc"
    End If
ElseIf mod1.KhK = 3 Then
    tt = "select * from bjdV where comid=" & comId & " order by 报价日期 desc"
End If
adoBJD.Close
adoBJD.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgBJD.DataSource = adoBJD
dtgBJD.Row = adoBJD.RecordCount - 1
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
If frmBxV.Visible = True Then
    frmBxV.Enabled = True
    frmBxV.ZOrder 0
Else
    frmZu.Enabled = True
    frmZu.ZOrder 0
End If
End Sub



Private Sub cmdCopy_Click()
dtgBB.FixedRows = 0
dtgBB.FixedCols = 0

    dtgBB.MergeCol(0) = False
    dtgBB.MergeCells = 0
dtgBB.Col = 0
dtgBB.Row = 0
If comLx.Text = "团队费用" Then
    dtgBB.ColSel = 11
ElseIf comLx.Text = "个人费用" Then
    dtgBB.ColSel = 23

ElseIf comLx.Text = "个人福利 提成" Then
    dtgBB.ColSel = 10
ElseIf comLx.Text = "公司费用明细" Then
    dtgBB.ColSel = 13
ElseIf comLx.Text = "应收帐款" Then
    dtgBB.ColSel = 6

End If
    dtgBB.RowSel = dtgBB.Rows - 3
Clipboard.Clear
Clipboard.SetText dtgBB.Clip
dtgBB.FixedRows = 1
If comLx.Text = "公司费用明细" Then
    dtgBB.FixedCols = 1
ElseIf comLx.Text = "应收帐款" Then
    dtgBB.FixedCols = 0
Else
    dtgBB.FixedCols = 2
End If
    dtgBB.MergeCol(0) = True
    dtgBB.MergeCells = 3
End Sub

Private Sub cmdJZ_Click()
Dim tt As String
On Error Resume Next
If mod1.DName <> "文静" Then
    Call mod1.NoQuan
    Exit Sub
End If
timZm = 1 '结帐添空
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "报销单"
    mod1.cmd.Parameters("@NBLX") = "结帐添空"
    mod1.cmd.Parameters("@bh") = ""
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
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
    mod1.cmd.Parameters("@md1") = dt2.Value
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

Private Sub cmdV_Click()
Dim tt As String
Dim ii As Integer
Dim OBm As String
Dim Nbm As String
Dim Ri As Integer
Dim oo As Integer
Dim FHg As Double
Dim Dhg As Double '当月合计
Dim Zhg As Double    '累积合计
Dim YBF As Boolean
Dim Ra: Dim ua
On Error Resume Next
'MsgBox "正在建设中！"
'Exit Sub
Me.Enabled = False
frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh
Me.MousePointer = 11
lblCell.Visible = False
If comLx.Text = "团队费用" Then
'''''''''''''''''    dtgBB.Clear
'''''''''''''''''    dtgBB.Row = 0: dtgBB.Col = 0: dtgBB.Text = "部门": dtgBB.Col = 1: dtgBB.Text = "姓名": dtgBB.Col = 2: dtgBB.Text = "当月办公用品": dtgBB.Col = 3: dtgBB.Text = "合计"
'''''''''''''''''    dtgBB.Col = 4: dtgBB.Text = "当月培训费": dtgBB.Col = 5: dtgBB.Text = "合计": dtgBB.Col = 6: dtgBB.Text = "当月团队建设费 ": dtgBB.Col = 7: dtgBB.Text = "合计"
'''''''''''''''''    dtgBB.Col = 8: dtgBB.Text = "当月固定资产": dtgBB.Col = 9: dtgBB.Text = "合计": dtgBB.Col = 10: dtgBB.Text = "当月合计": dtgBB.Col = 11: dtgBB.Text = "合计"
    
    tt = "select bm as 部门,姓名,sum(办公用品) as 当月办公用品,(select sum(办公用品) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    "sum(培训费) as 当月培训费,(select sum(培训费) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    "sum(团队建设费)+sum(部门团队费) as 当月团队建设费,(select sum(团队建设费)+sum(部门团队费) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    " '' as 当月固定资产,'' as 合计," & _
    " sum(办公用品+培训费+团队建设费+部门团队费) as 当月合计," & _
    " (select sum(办公用品+培训费+团队建设费+部门团队费) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计" & _
    " ,ywyuid,max(bmid) as bmid from 费用统计A as P where" & _
        " year(日期)=" & Year(dt2.Value) & " and month(日期)=" & Month(dt2.Value) & "   group by bm,bmid,姓名,ywyuid  order by bmid"
ElseIf comLx.Text = "个人费用" Then
'''''        " year(日期)=" & Year(dt2.Value) & " and month(日期)=" & Month(dt2.Value) & _
'''''            " and (bm='维销部1' or bm='维销部2' or bm='北京办' or bm='南京办' or bm='杭州办' or bm='产品部1' or bm='产品部2')   group by bm,姓名,qy,ywyuid  order by qy,bm"
    tt = "select bm as 部门,姓名,sum(办公用品) as 当月办公用品,(select sum(办公用品) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    "sum(通信费) as 当月通信费,(select sum(通信费) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    "sum(市内交通费) as 当月市内交通费,(select sum(市内交通费) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    "sum(市外交通费) as 当月市外交通费,(select sum(市外交通费) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    "sum(住宿费) as 当月住宿费,(select sum(住宿费) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    "sum(招待费) as 当月招待费,(select sum(招待费) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    "sum(礼品费) as 当月礼品费,(select sum(礼品费) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    "sum(餐费) as 当月餐费,(select sum(餐费) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    "sum(运费)+sum(快递费) as 当月运费,(select sum(运费)+sum(快递费) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    "sum(财务手续费) as 当月财务手续费,(select sum(财务手续费)  from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    " sum(办公用品+通信费+市内交通费+市外交通费+住宿费+招待费+礼品费+餐费+运费+快递费+财务手续费) as 当月合计," & _
    " (select sum(办公用品+通信费+市内交通费+市外交通费+住宿费+招待费+礼品费+餐费+运费+快递费+财务手续费) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计" & _
    " ,ywyuid,max(bmid) as bmid from 费用统计A as P where" & _
        " year(日期)=" & Year(dt2.Value) & " and month(日期)=" & Month(dt2.Value) & "   group by bm,bmid,姓名,ywyuid  order by bmid"
            
ElseIf comLx.Text = "个人福利 提成" Then
    tt = "select bm as 部门,姓名,'' as 当月工资,'' as 合计," & _
    "sum(四金)+sum(房屋补贴)+sum(旅游费)+sum(通信费)+sum(高温费)+sum(福利费)+sum(驻外津贴)+sum(交通补贴)+sum(岗位补贴)+sum(公积金)+sum(综合保险) as 当月福利," & _
    "(select sum(四金)+sum(房屋补贴)+sum(旅游费)+sum(通信费)+sum(高温费)+sum(福利费)+sum(驻外津贴)+sum(交通补贴)+sum(岗位补贴)+sum(公积金)+sum(综合保险) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    "sum(四金)+sum(房屋补贴)+sum(旅游费)+sum(通信费)+sum(高温费)+sum(福利费)+sum(驻外津贴)+sum(交通补贴)+sum(岗位补贴)+sum(公积金)+sum(综合保险) as 当月合计," & _
    "(select sum(四金)+sum(房屋补贴)+sum(旅游费)+sum(通信费)+sum(高温费)+sum(福利费)+sum(驻外津贴)+sum(交通补贴)+sum(岗位补贴)+sum(公积金)+sum(综合保险) from 费用统计A where ywyuid=p.ywyuid and bmid=P.bmid  and 日期>='" & _
    dt1.Value & "' and 日期<'" & DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1) & "') as 合计," & _
    " ''as 当月提成,'' as 合计,'' as 当月奖金,'' as 合计,'' as 当月合计,'' as 合计,ywyuid,max(bmid) as bmid from 费用统计A as P where" & _
        " year(日期)=" & Year(dt2.Value) & " and month(日期)=" & Month(dt2.Value) & "   group by bm,bmid,姓名,ywyuid  order by bmid"
ElseIf comLx.Text = "公司费用明细" Then
    tt = "SELECT month(日期) as 日期,sum(房租) as 房租,sum(水电) as 水电,sum(物业费) as 物业费,sum(电话) as 电话,0 as 固定资产,sum(人员招聘) as 人员招聘," & _
          "sum(市场推广) as 市场推广,sum(公共停车费+公共车辆费) as 公共车辆费,0 as 税金,sum(房租+水电+物业费+电话+人员招聘+市场推广+公共停车费+公共车辆费) as 合计" & _
      " FROM  费用统计A  where 日期>='" & dt1.Value & "' and 日期<'" & dt2.Value & "' group by year(日期),month(日期) order by year(日期),month(日期)"
End If

dtgBB.FixedCols = 0
Set adoBB = New ADODB.Recordset
adoBB.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
dtgBB.FixedRows = 1
dtgBB.Visible = False
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
ua = UBound(Ra, 2)
For oo = 1 To ua + 1
    dtgBB.Row = oo
    For ii = 0 To 20
        dtgBB.Col = ii
        dtgBB.Text = Trim(Ra(ii, oo - 1))
    Next
Next


If comLx.Text = "公司费用明细" Then
    If Month(dt1.Value) >= 4 Then
        dt1.Value = DateSerial(Year(dt1.Value), 4, 1)
        dt2.Value = DateSerial(Year(dt1.Value) + 1, 3, 31)
    Else
        dt1.Value = DateSerial(Year(dt1.Value) - 1, 4, 1)
        dt2.Value = DateSerial(Year(dt1.Value), 3, 31)
    End If
    frmWait.Visible = False
    Me.Enabled = True
    dtgBB.Visible = False
    dtgBB.FixedRows = 0
    
    Set dtgBB.DataSource = Nothing
    dtgBB.Clear
    Set dtgCol.DataSource = adoBB
    '行列转换
    For oo = 0 To dtgCol.Rows - 1
        dtgBB.Col = oo
        dtgCol.Row = oo
        For ii = 0 To dtgCol.Cols - 1
            dtgBB.Row = ii
            dtgCol.Col = ii
            dtgBB.Text = dtgCol.Text
        Next
    Next
    dtgBB.FixedRows = 1
    dtgBB.FixedCols = 1
    '设置格式
    For oo = 0 To 12
        dtgBB.Col = oo
        For ii = 0 To 12
            dtgBB.Row = ii
            If ii = 0 And oo = 0 Then
                dtgBB.Text = "明细"
            End If
            If ii = 0 Then
                If oo > 0 Then
                    dtgBB.Text = dtgBB.Text & "月"
                End If
                dtgBB.CellFontBold = True
            End If
            If Val(dtgBB.Text) = 0 And dtgBB.Col <> 0 And dtgBB.Row <> 10 And dtgBB.Row <> 11 Then
                dtgBB.Text = ""
            End If
        Next
    Next
    dtgBB.Row = 10: Zhg = 0
    For oo = 0 To 12 '设置合计颜色
        Zhg = Round(Zhg + Val(dtgBB.Text), 2)
        dtgBB.Col = oo
        dtgBB.CellBackColor = &HFF&
        dtgBB.CellFontBold = True
    Next
    dtgBB.Row = 11: dtgBB.Col = 0
    dtgBB.Text = "总计": dtgBB.CellForeColor = &HFF&: dtgBB.CellFontBold = True
    dtgBB.Col = 1: dtgBB.Text = Zhg: dtgBB.CellForeColor = &HFF&: dtgBB.CellFontBold = True
    dtgBB.Visible = True
    Me.ZOrder 0
    Me.MousePointer = 0
    Exit Sub
End If

dtgBB.Visible = False
Set dtgBB.DataSource = adoBB
dtgBB.Rows = dtgBB.Rows + 50

Ri = 2
dtgBB.Col = 0
dtgBB.Row = 1
OBm = Trim(dtgBB.Text)
Nbm = ""
For oo = 2 To dtgBB.Rows - 1
    dtgBB.Row = Ri

    If OBm <> Trim(dtgBB.Text) Then
        Nbm = Trim(dtgBB.Text)
        dtgBB.AddItem OBm & "合计", Ri
        dtgBB.CellBackColor = &HFF&
        dtgBB.Col = 1
        dtgBB.CellBackColor = &HFF&
        dtgBB.Col = 0

        dtgBB.AddItem "", Ri + 1
        OBm = Nbm
        Ri = Ri + 2
        oo = oo + 2
    
    End If
    Ri = Ri + 1
'''''        If Trim(dtgBB.Text) = "总经理" Then
'''''            dtgBB.Col = 0
'''''            dtgBB.CellBackColor = &H8000000D
'''''            dtgBB.Col = 1
'''''            dtgBB.CellBackColor = &H8000000D
'''''            dtgBB.Col = 0
'''''        End If
Next

'计算合计
dtgBB.Col = 2
dtgBB.Row = 1
FHg = 0
Dhg = 0
Zhg = 0
For ii = 2 To dtgBB.Cols - 3
    dtgBB.Col = ii
    For oo = 1 To dtgBB.Rows + 50
        dtgBB.Row = oo

        dtgBB.Col = 0
        If Right(dtgBB.Text, 2) <> "合计" Then
            dtgBB.Col = ii
            FHg = FHg + Val(dtgBB.Text)
        Else
            dtgBB.Col = ii
            dtgBB.Text = FHg
            Zhg = Zhg + FHg
            FHg = 0
            dtgBB.CellBackColor = &HFF&
        End If
        If Val(dtgBB.Text) = 0 Then
            dtgBB.Col = 0
            If Right(dtgBB.Text, 2) <> "合计" Then
                dtgBB.Col = ii
                dtgBB.Text = ""
                
            End If
            dtgBB.Col = ii
        End If
    Next
Next
If adoBB.RecordCount > 0 Then
    dtgBB.FixedRows = 0
    dtgBB.MergeCol(0) = True
    dtgBB.MergeCells = 3
    dtgBB.FixedRows = 1
End If
dtgBB.FixedCols = 2
dtgBB.Col = 0
dtgBB.Row = 0
dtgBB.CellFontBold = True
dtgBB.Col = 1
dtgBB.CellFontBold = True
For ii = 2 To dtgBB.Cols - 1
    dtgBB.Col = ii
    YBF = False
    dtgBB.Row = 0
    dtgBB.CellFontBold = True
    If Left(dtgBB.Text, 2) <> "当月" Then
        YBF = True
    End If
    For oo = 0 To dtgBB.Rows + 50
        dtgBB.Row = oo
        If YBF = True Then
            dtgBB.CellForeColor = &H8000000D
            dtgBB.CellFontBold = True
        End If
    Next
Next
If comLx.Text = "团队费用" Then
    dtgBB.ColWidth(12) = 0
    dtgBB.ColWidth(13) = 0
    dtgBB.Cols = dtgBB.Cols + 10
ElseIf comLx.Text = "个人费用" Then
    dtgBB.ColWidth(24) = 0
    dtgBB.ColWidth(25) = 0
ElseIf comLx.Text = "个人福利 提成" Then
    dtgBB.ColWidth(14) = 0
    dtgBB.ColWidth(15) = 0
    dtgBB.Cols = dtgBB.Cols + 10
End If
dtgBB.Visible = True

'''dtgBB.Row = 0
'''dtgBB.CellFontBold = True
Me.Enabled = True
frmWait.Visible = False
Me.ZOrder 0
Me.MousePointer = 0
End Sub

Private Sub cmdXuan_Click()
dtgBB.FixedRows = 0
dtgBB.FixedCols = 0

    dtgBB.MergeCol(0) = False
    dtgBB.MergeCells = 0
dtgBB.Col = 0
dtgBB.Row = 0
dtgBB.ColSel = 23
dtgBB.RowSel = dtgBB.Rows - 3
End Sub

Private Sub cmdVnew_Click()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
Dim FHg As Double
Dim Obmid As Integer
Dim Ra: Dim ua: Dim Rb: Dim ub
Dim oClo
On Error Resume Next
If comLx.Enabled = True Then
'''    tt = "select * from Fk where year(rq)=" & Year(dt2.Value) & " and month(rq)=" & Month(dt2.Value) & " order by bmid,xuid,rq;" & _
'''        "select round(sum(yingfje),2) from Fk where year(rq)=" & Year(dt2.Value) & " and month(rq)=" & Month(dt2.Value)
    tt = "select * from Fk where rq>='" & DateSerial(Year(dt1.Value), Month(dt1.Value) - 1, 26) & "' and rq<'" & DateSerial(Year(dt2.Value), Month(dt2.Value), 26) & "' order by bmid,xuid,rq;" & _
        "select round(sum(yingfje),2) from Fk where rq>='" & DateSerial(Year(dt1.Value), Month(dt1.Value) - 1, 26) & "' and rq<'" & DateSerial(Year(dt2.Value), Month(dt2.Value), 26) & "'"
Else
    tt = "select * from Fk where rq>='" & DateSerial(Year(dt1.Value), Month(dt1.Value) - 1, 26) & "' and rq<'" & DateSerial(Year(dt2.Value), Month(dt2.Value), 26) & "' and ggl='" & mod1.DHid & "' order by bmid,xuid,rq;" & _
        "select round(sum(yingfje),2) from Fk where rq>='" & DateSerial(Year(dt1.Value), Month(dt1.Value) - 1, 26) & "' and rq<'" & DateSerial(Year(dt2.Value), Month(dt2.Value), 26) & "' and ggl='" & mod1.DHid & "'"
End If
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
dtgBB.Visible = False
dtgBB.Clear
dtgBB.Cols = 18
dtgBB.Row = 0: dtgBB.Col = 0: dtgBB.Text = "区域": dtgBB.Col = 1: dtgBB.Text = "部门": dtgBB.Col = 2: dtgBB.Text = "业务员": dtgBB.Col = 3: dtgBB.Text = "应收日期"
dtgBB.Col = 4: dtgBB.Text = "金额": dtgBB.Col = 5: dtgBB.Text = "合同编号": dtgBB.Col = 6: dtgBB.Text = "项目名称"
dtgBB.FixedCols = 0
For ii = 0 To 12
    dtgBB.CellFontBold = True
    dtgBB.Col = ii
    dtgBB.ColWidth(ii) = 0
    If dtgBB.Text = "应收日期" Then
        dtgBB.ColWidth(ii) = 1500
    End If
    If dtgBB.Text = "金额" Then
        dtgBB.ColWidth(ii) = 1500
    End If
    If dtgBB.Text = "合同编号" Then
        dtgBB.ColWidth(ii) = 2000
    End If
    If dtgBB.Text = "项目名称" Then
        dtgBB.ColWidth(ii) = 3000
    End If
    If dtgBB.Text = "区域" Or dtgBB.Text = "部门" Or dtgBB.Text = "业务员" Then
        dtgBB.ColWidth(ii) = 1000
    End If
Next
ua = UBound(Ra, 2)
dtgBB.Rows = ua + 50
For oo = 1 To ua + 1
    dtgBB.Row = oo
    For ii = 0 To 12
        dtgBB.Col = ii
        dtgBB.Text = Trim(Ra(ii, oo - 1))
    Next
Next
Obmid = -1: oClo = &HC0FFC0
For oo = 1 To ua + 1
    dtgBB.Row = oo
    dtgBB.Col = 7
    If Obmid <> Val(dtgBB.Text) Then
        Obmid = Val(dtgBB.Text)
        If oClo = &HC0FFC0 Then

            oClo = &HC0FFFF
        Else
            oClo = &HC0FFC0
        End If
        For ii = 0 To 9
            dtgBB.Col = ii
                dtgBB.CellBackColor = oClo
        Next
    Else
        For ii = 0 To 9
                dtgBB.Col = ii
                dtgBB.CellBackColor = oClo
        Next
    End If
    oClo = dtgBB.CellBackColor
Next
dtgBB.FixedRows = 1
dtgBB.Visible = True

dtgBB.Row = ua + 3
dtgBB.Col = 3
dtgBB.Text = "当月预计应收款："
dtgBB.CellFontBold = True
dtgBB.CellForeColor = &HFF&
dtgBB.Col = 4
dtgBB.Text = Rb(0, 0)
dtgBB.CellFontBold = True
dtgBB.CellForeColor = &HFF&
    dtgBB.MergeCol(0) = True
    dtgBB.MergeCol(1) = True
    dtgBB.MergeCol(2) = True
    dtgBB.MergeCells = 3
End Sub

Private Sub Command1_Click()

End Sub

Private Sub comLx_Click()
If comLx.Text = "应收帐款" Then
    cmdVnew.Visible = True: cmdV.Visible = False
Else
    cmdV.Visible = True: cmdVnew.Visible = False
End If

End Sub

Private Sub dt1_CloseUp()
dt2.Value = dt1.Value
End Sub

Private Sub dtgBB_Click()




'''''dtgBB.CellBackColor = &HFFFF&
End Sub

Private Sub dtgBB_DblClick()
Dim Orow As Integer
Dim OCol As Integer
Dim Bmid As Integer
Dim YwyUid As String
Dim LM As String '列名
Dim FR As Date
Dim LR As Date
Dim tt As String
Dim hg As Double
Dim oo As Integer
Dim Hid As Long
Dim NewF As Integer
Dim MM As Integer '月份
On Error Resume Next

If comLx.Text = "应收帐款" Then
    mod1.BTZ = 6
    dtgBB.Col = 8
    Hid = Val(dtgBB.Text)
    dtgBB.Col = 10
    NewF = Val(dtgBB.Text)
    If NewF = 1 Then
            Call modHt.NewQing
            Call modHt.NewLocked
            Call modHt.NewBound(Hid)
            frmWbNew.Visible = True
            frmWbNew.ZOrder 0
    ElseIf NewF = 2 Then
            Call modNewHT.NewMQing
    
            Call modNewHT.NewMBound(Hid)
            FMXC.lblMQM(0).Visible = True
            FMXC.lblMTm(0).Visible = True
            FMXC.cmdMQm(0).Visible = True
            FMXC.ZOrder 0
    ElseIf NewF >= 3 Then
            Call modNewHT.NewMQing
            
            Call modNewHT.NewB(Hid)
            FMXC.lblMQM(0).Visible = True
            FMXC.lblMTm(0).Visible = True
            FMXC.cmdMQm(0).Visible = True
            FMXC.ZOrder 0
    Else
        MsgBox "此为旧版合同，请通过合同列表查询！"
        Exit Sub
    End If
    Me.Enabled = False
    Exit Sub
End If
YwyUid = ""
Bmid = 0
If Val(dtgBB.Text) = 0 Then
    Exit Sub
End If
Orow = dtgBB.Row
OCol = dtgBB.Col
If comLx.Text <> "公司费用明细" Then
    dtgBB.Row = 0
    '取得列名
    If dtgBB.Text = "合计" Then
        dtgBB.Col = dtgBB.Col - 1
        LM = Mid(dtgBB.Text, 3, Len(dtgBB.Text) - 1)
        dtgCol = dtgCol + 1
        FR = DateSerial(Year(dt1.Value), Month(dt1.Value), 1)
        LR = DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1)
    Else
        LM = Mid(dtgBB.Text, 3, Len(dtgBB.Text) - 1)
        FR = DateSerial(Year(dt2.Value), Month(dt2.Value), 1)
        LR = DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1)
    End If
    dtgBB.Row = Orow
    If comLx.Text = "团队费用" Then
        dtgBB.Col = 12
        YwyUid = Trim(dtgBB.Text)
        dtgBB.Col = 13
        Bmid = Val(dtgBB.Text)
    ElseIf comLx.Text = "个人费用" Then
        dtgBB.Col = 24
        YwyUid = Trim(dtgBB.Text)
        dtgBB.Col = 25
        Bmid = Val(dtgBB.Text)
        
    ElseIf comLx.Text = "个人福利 提成" Then
        dtgBB.Col = 14
        YwyUid = Trim(dtgBB.Text)
        dtgBB.Col = 15
        Bmid = Val(dtgBB.Text)
    End If
    dtgBB.Col = OCol
    If YwyUid = "" Or Bmid = 0 Or LM = "合计" Then
        Exit Sub
    End If
    tt = "select 日期 as 报销日期,内容 as 报销内容," & LM & ",bxid as 编号 from 费用统计A where ywyuid='" & YwyUid & "' and bmid=" & Bmid & " and 日期>='" & FR & "' and 日期<'" & LR & "' and " & LM & ">0 order by 日期 desc"
        frmCWBBA.BCol = 3
    If LM = "团队建设费" Then
        tt = "select 日期 as 报销日期,内容 as 报销内容," & LM & ",部门团队费,bxid as 编号 from 费用统计A where ywyuid='" & YwyUid & "' and bmid=" & _
        Bmid & " and 日期>='" & FR & "' and 日期<'" & LR & "' and (团队建设费>0 or 部门团队费>0) order by 日期 desc"
        frmCWBBA.BCol = 4
    ElseIf LM = "运费" Then
        tt = "select 日期 as 报销日期,内容 as 报销内容," & LM & ",快递费,bxid as 编号 from 费用统计A where ywyuid='" & YwyUid & "' and bmid=" & _
        Bmid & " and 日期>='" & FR & "' and 日期<'" & LR & "' and (运费>0 or 快递费>0) order by 日期 desc"
        frmCWBBA.BCol = 4
    ElseIf LM = "福利" Then
        tt = "select 日期 as 报销日期,内容 as 报销内容,四金 as 三金,房屋补贴,旅游费,通信费,高温费,福利费,驻外津贴,交通补贴,岗位补贴,公积金,综合保险 ,bxid as 编号 from 费用统计A where ywyuid='" & YwyUid & "' and bmid=" & _
        Bmid & " and 日期>='" & FR & _
        "' and 日期<'" & LR & "' and (四金>0 or 房屋补贴>0 or 旅游费>0 or 通信费>0 or 高温费>0 or 福利费>0 or 驻外津贴>0 or 交通补贴>0 or 岗位补贴>0 or 公积金>0 or 综合保险>0) order by 日期 desc"
        frmCWBBA.BCol = 13
    End If
Else '
    FR = DateSerial(Year(dt1.Value), Month(dt1.Value), 1)
    LR = DateSerial(Year(dt2.Value), Month(dt2.Value) + 1, 1)
    dtgBB.Col = 0
    LM = Trim(dtgBB.Text)
    dtgBB.Col = OCol
    dtgBB.Row = 0
    MM = Val(dtgBB.Text)
    dtgBB.Row = Orow
    If LM = "合计" Or MM = 0 Or LM = "总计" Then
        Exit Sub
    End If
    If LM = "公共车辆费" Then
        tt = "select 日期 as 报销日期,内容 as 报销内容,公共车辆费,公共停车费,bxid as 编号 from 费用统计A where  日期>='" & _
        FR & "' and 日期<'" & LR & "' and (公共车辆费>0 or 公共停车费>0) and month(日期)=" & MM & "  order by 日期 desc"
        frmCWBBA.BCol = 4
    Else
        tt = "select 日期 as 报销日期,内容 as 报销内容," & LM & ",bxid as 编号 from 费用统计A where  日期>='" & _
        FR & "' and 日期<'" & LR & "' and " & LM & ">0 and month(日期)=" & MM & "  order by 日期 desc"
        frmCWBBA.BCol = 3
    End If

End If
Set frmCWBBA.adoL = New ADODB.Recordset
frmCWBBA.adoL.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText

If frmCWBBA.adoL.RecordCount = 0 Then
    frmCWBBA.dtgL.FixedRows = 0
End If
Set frmCWBBA.dtgL.DataSource = frmCWBBA.adoL
    frmCWBBA.dtgL.FixedRows = 1
frmCWBBA.dtgL.Rows = frmCWBBA.dtgL.Rows + 30
frmCWBBA.dtgL.Cols = frmCWBBA.dtgL.Cols + 5
hg = 0
frmCWBBA.dtgL.Row = 1
frmCWBBA.dtgL.Col = 2
oo = 1
Do While Not frmCWBBA.dtgL.Col > frmCWBBA.BCol - 1
    Do While Not oo > frmCWBBA.adoL.RecordCount
        hg = hg + Val(frmCWBBA.dtgL.Text)
        'frmCWBBA.adoL.MoveNext

        oo = oo + 1
        frmCWBBA.dtgL.Row = oo
    Loop
    frmCWBBA.dtgL.Col = frmCWBBA.dtgL.Col + 1
    oo = 1
    frmCWBBA.dtgL.Row = oo
Loop
frmCWBBA.dtgL.Row = frmCWBBA.adoL.RecordCount + 1
frmCWBBA.dtgL.Col = 2
frmCWBBA.dtgL.Text = hg
frmCWBBA.dtgL.CellFontBold = True
frmCWBBA.dtgL.Col = 1
frmCWBBA.dtgL.Text = "合计"
frmCWBBA.dtgL.CellFontBold = True
frmCWBBA.Show
Me.Enabled = False
frmCWBBA.ZOrder 0
End Sub


Private Sub dtgBB_LeaveCell()
'''''Orow = dtgBB.Row
'''''Ocol = dtgBB.Col
'''''Oc = dtgBB.CellBackColor
End Sub

Private Sub dtgBB_RowColChange()
'''''dtgBB.Col = Ocol
'''''dtgBB.Row = Orow
'''''dtgBB.CellBackColor = Oc
End Sub

Private Sub Form_Load()
On Error Resume Next
Set adoBJD = New ADODB.Recordset
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
dtgBJD.ColWidth(0) = 300
dtgBJD.ColWidth(1) = 3500
dtgBJD.ColWidth(4) = 0
dtgBJD.ColWidth(6) = 0
dtgBJD.ColWidth(7) = 0
dtgBJD.ColWidth(8) = 0
If mod1.Mname = "马晓聪" Then
    frmNew.Visible = True
Else
    frmNew.Visible = False
End If
dt1.Value = DateSerial(Year(Date) - 1, 4, 1)
dt2.Value = Date
Me.Left = 0
Me.Top = 0
dtgXX.Left = dtgBB.Left
dtgXX.Top = dtgBB.Top
dtgXX.Visible = False
dtgBB.Visible = True
dtgXX.Cols = dtgBB.Cols
dtgBB.Rows = dtgBB.Rows + 50
cmdVnew.Left = cmdV.Left
cmdVnew.Top = cmdV.Top
cmdVnew.Visible = False: cmdV.Visible = True
End Sub


Private Sub OKButton_Click()
Dim tt As String
Dim ZL As String
Dim BaoId As Long
On Error Resume Next
dtgBJD.Col = 6
BaoId = dtgBJD.Text
dtgBJD.Col = 7
ZL = dtgBJD.Text
Call modBJD.BaoJDBound(BaoId, ZL)
frmWbxjB.cmdSave.Enabled = False
frmWbxjB.cmdMod.Enabled = False
frmGxbjB.cmdSave.Enabled = False
frmGxbjB.cmdMod.Enabled = False
End Sub


Private Sub timQuit_Timer()
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0

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
'    Unload frmWaitA
'    Me.Enabled = True
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    timWait.Enabled = False
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0

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


