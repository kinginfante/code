VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form fmxcY 
   BackColor       =   &H00C0FFC0&
   Caption         =   "财务到帐"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   10320
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgXLN 
      Height          =   255
      Left            =   5010
      TabIndex        =   44
      Top             =   5580
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgFKN 
      Height          =   525
      Left            =   9630
      TabIndex        =   43
      Top             =   5130
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   926
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame frmXX 
      BackColor       =   &H00C0FFC0&
      Height          =   1635
      Left            =   1410
      TabIndex        =   31
      Top             =   3210
      Width           =   8775
      Begin VB.CommandButton cmdD 
         Caption         =   "删除"
         Height          =   285
         Left            =   3270
         TabIndex        =   45
         Top             =   570
         Width           =   675
      End
      Begin VB.CommandButton cmdGx 
         Caption         =   "更新"
         Height          =   270
         Left            =   3270
         TabIndex        =   40
         Top             =   960
         Width           =   675
      End
      Begin VB.TextBox txtXrq 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   930
         Width           =   2115
      End
      Begin VB.TextBox txtJJ 
         Height          =   270
         Left            =   720
         TabIndex        =   36
         Top             =   540
         Width           =   2385
      End
      Begin VB.CommandButton cmdC 
         Caption         =   "查询"
         Height          =   285
         Left            =   3270
         TabIndex        =   34
         Top             =   120
         Width           =   705
      End
      Begin VB.TextBox txtZ 
         Height          =   270
         Left            =   720
         TabIndex        =   33
         Top             =   120
         Width           =   2355
      End
      Begin MSComCtl2.DTPicker dtgXrq 
         Height          =   315
         Left            =   720
         TabIndex        =   38
         Top             =   930
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   8454016
         CalendarTitleBackColor=   16711808
         CalendarTrailingForeColor=   -2147483635
         Format          =   109969409
         CurrentDate     =   38797
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgFk 
         Height          =   1665
         Left            =   4110
         TabIndex        =   42
         Top             =   0
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   2937
         _Version        =   393216
         BackColor       =   16777152
         Rows            =   30
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   12648384
         BackColorBkg    =   16777152
         FillStyle       =   1
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "日期"
         Height          =   225
         Left            =   120
         TabIndex        =   39
         Top             =   1020
         Width           =   465
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "金额"
         Height          =   255
         Left            =   90
         TabIndex        =   35
         Top             =   570
         Width           =   465
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "编号"
         Height          =   225
         Left            =   90
         TabIndex        =   32
         Top             =   150
         Width           =   435
      End
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1950
      Top             =   0
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   30
      Top             =   450
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "添加"
      Height          =   585
      Left            =   6540
      Picture         =   "fmxcY.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4920
      Width           =   675
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除"
      Enabled         =   0   'False
      Height          =   585
      Left            =   8580
      Picture         =   "fmxcY.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4920
      Width           =   645
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   585
      Left            =   7260
      Picture         =   "fmxcY.frx":05CC
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4920
      Width           =   645
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   9240
      Picture         =   "fmxcY.frx":0A0E
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4920
      Width           =   585
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "提交"
      Height          =   585
      Left            =   7890
      Picture         =   "fmxcY.frx":0B10
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4920
      Width           =   675
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   0
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   5085
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1305
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   300
         Width           =   4125
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   4260
         TabIndex        =   13
         Top             =   510
         Width           =   705
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   4260
         TabIndex        =   12
         Top             =   870
         Width           =   675
      End
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   4290
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "提交审核"
      Height          =   585
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   645
   End
   Begin VB.Frame frmYwy 
      BackColor       =   &H00C0FFC0&
      Caption         =   "业务助理填写(注明合同编号)"
      ForeColor       =   &H00000000&
      Height          =   3075
      Left            =   4260
      TabIndex        =   4
      Top             =   90
      Width           =   6045
      Begin VB.TextBox txtBz 
         Height          =   795
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "fmxcY.frx":117A
         Top             =   300
         Width           =   5895
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgXL 
         Height          =   1635
         Left            =   120
         TabIndex        =   41
         Top             =   1260
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   2884
         _Version        =   393216
         BackColor       =   12648384
         BackColorFixed  =   12648384
         BackColorBkg    =   12648384
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame frmCw 
      BackColor       =   &H00C0FFC0&
      Caption         =   "财务填写"
      ForeColor       =   &H00000000&
      Height          =   3075
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   4245
      Begin VB.ComboBox companyId 
         Height          =   300
         ItemData        =   "fmxcY.frx":1180
         Left            =   1080
         List            =   "fmxcY.frx":118D
         TabIndex        =   29
         Text            =   "上海豪曼制冷空调服务有限公司"
         Top             =   2190
         Width           =   3015
      End
      Begin VB.ComboBox comQy 
         Height          =   300
         ItemData        =   "fmxcY.frx":11E3
         Left            =   3060
         List            =   "fmxcY.frx":11F6
         TabIndex        =   27
         Text            =   "上海"
         Top             =   1680
         Width           =   1035
      End
      Begin VB.TextBox txtDate 
         Height          =   270
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1110
         Width           =   2745
      End
      Begin VB.TextBox txtJe 
         Height          =   315
         Left            =   1080
         TabIndex        =   8
         Text            =   "Text3"
         Top             =   1650
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   1110
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Format          =   141688833
         CurrentDate     =   40476
      End
      Begin VB.TextBox txtKhmc 
         Height          =   645
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "fmxcY.frx":1218
         Top             =   300
         Width           =   2985
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "公    司  "
         Height          =   225
         Left            =   150
         TabIndex        =   30
         Top             =   2250
         Width           =   795
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "区域"
         Height          =   225
         Left            =   2520
         TabIndex        =   26
         Top             =   1740
         Width           =   465
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "到帐金额"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   1740
         Width           =   825
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "到帐日期"
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   1170
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "客户名称"
         Height          =   405
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   885
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   1695
      Left            =   0
      TabIndex        =   15
      Top             =   3780
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   2990
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
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
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
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblTX 
      BackStyle       =   0  'Transparent
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
      Left            =   5910
      TabIndex        =   25
      Top             =   3900
      Width           =   4815
   End
   Begin VB.Label lblAid 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      Height          =   165
      Left            =   8940
      TabIndex        =   24
      Top             =   3510
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ID号:"
      Height          =   195
      Left            =   8400
      TabIndex        =   23
      Top             =   3510
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "评审状态"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   90
      TabIndex        =   16
      Top             =   3330
      Width           =   1005
   End
End
Attribute VB_Name = "fmxcY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer '成本追加单(1保存2删除3明细编辑5签字19通知执行)

Dim LCRen As String
Dim LCUid As String
Public Lc As Integer
Dim Fwid As Long
Public Ywy As String
Public Uid As String
Dim Jid As Integer


Public Sub QMBound(Zid As Long)
Dim Ra: Dim La
Dim ii As Integer: Dim oo As Integer
Dim tt As String
On Error Resume Next

tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & Zid & "' and yid=91 order by pid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2): dtgP.Rows = La + 20
Call dtgPFF
For oo = 1 To La + 1
    dtgP.Row = oo
    For ii = 0 To 5
        dtgP.Col = ii
        dtgP.Text = Ra(ii, oo - 1)
            DH = 255 * mod1.HH(dtgP.Text, UpInt(dtgP.CellWidth / 100))
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



End Sub
Public Sub Qing()
txtKhmc.Text = ""
txtDate.Text = ""
txtJe.Text = ""
txtBz.Text = ""
lblAid.Caption = ""
Call dtgPFF

txtKhmc.Locked = True
dtpDate.Enabled = False
txtJe.Locked = True
comQy.Locked = True
txtBz.Locked = True
comQy.Text = ""
cmdSave.Enabled = False
cmdDel.Enabled = False
lblTX.Caption = ""

Lc = 0
LCRen = ""
LCUid = ""
Fwid = 0
Me.Ywy = ""
Me.Uid = ""
companyId.Text = "上海豪曼制冷空调服务有限公司"
frmXX.Visible = False
txtJJ.Text = ""
txtZ.Text = ""
txtXrq.Text = ""
Call Me.dtgXlFF
Call Me.FKFF
End Sub

Private Sub cmdAdd_Click()
If mod1.Bm <> "商务部" Then Exit Sub
Call Qing
    txtKhmc.Locked = False
    dtpDate.Enabled = True
    txtJe.Locked = False
    comQy.Locked = False
    cmdSave.Enabled = True
End Sub

Public Sub dtgPFF()
Dim oo As Integer
For oo = 1 To dtgP.Rows - 1
    dtgP.RowHeight(oo) = dtgP.RowHeight(0)
Next
dtgP.Clear
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "日期": dtgP.Col = 1: dtgP.Text = "姓名": dtgP.Col = 2: dtgP.Text = "职能": dtgP.Col = 3: dtgP.Text = "评审建议": dtgP.Col = 4: dtgP.Text = "审核":
dtgP.ColWidth(0) = 1005
dtgP.ColWidth(1) = 1005
dtgP.ColWidth(2) = 0
 dtgP.ColWidth(3) = 2310: dtgP.ColWidth(4) = 1035
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf FMXCYBR.Visible = True Then
    Call FMXCYBR.REF(FMXCYBR.tt)
End If
End Sub

Private Sub cmdC_Click()
Dim tt As String
Dim Rb, Ra
Dim HTZE As Single
Dim oo As Integer
Dim Lb As Integer
tt = "select 应付日期,收款额度,应付金额,fid,kdfh from htFK where right(htbh,5)='" & Right(txtZ.Text, 5) & "';" & _
    "select htze from htping where hid=" & Val(Right(txtZ.Text, 5))
'tt = "select rq,yingfJe,hxrq,hx,fid from htping1 where hid=" & Val(Right(txtZ.Text, 5))
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Lb = UBound(Rb, 2) + 1
HTZE = Ra(0, 0)
Call Me.FKBound(Rb, Lb, HTZE)
End Sub

Private Sub cmdD_Click()
Dim ii As Integer
If Jid = 0 Then Exit Sub
If txtXrq.Text = "" Then Exit Sub
ii = MsgBox("是否删除此笔结算记录？", vbYesNo + vbQuestion, "请确认！")
If ii = vbNo Then Exit Sub

frmQm.Visible = False
        timZm = 8 '删除结算
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "财务到帐"
        mod1.cmd.Parameters("@NBLX") = "删除结算"
        mod1.cmd.Parameters("@bh") = lblAid.Caption
       
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = ""
        mod1.cmd.Parameters("@mt2") = ""
        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1").Value = Jid
        mod1.cmd.Parameters("@mm2").Value = Val(txtJJ.Text)
        mod1.cmd.Parameters("@mb1") = 0

        mod1.cmd.Parameters("@md1") = txtXrq.Text
 
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

Private Sub cmdDel_Click()
Dim ii As Integer

ii = MsgBox("是否删除此财务到帐单?", vbYesNo + vbQuestion, "请确认")
If ii = vbNo Then Exit Sub

timZm = 2 '删除
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "财务到帐"
    mod1.cmd.Parameters("@NBLX") = "删除"
    mod1.cmd.Parameters("@bh") = lblAid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = ""
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""

    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Fwid

        mod1.cmd.Parameters("@mb1") = 0

    mod1.cmd.Parameters("@md1") = Null

    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 2 Then '保存
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
End Sub

Private Sub cmdDing_Click()
On Error Resume Next
If Lc = 0 Then
    Exit Sub
End If
If comQy.Text = "" Then
    MsgBox "请确认业务归属!"
    Exit Sub
End If
If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If

frmQm.Visible = False
        timZm = 5 '签字
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "财务到帐"
        mod1.cmd.Parameters("@NBLX") = "签字"
        mod1.cmd.Parameters("@bh") = lblAid.Caption
        If mod1.cmd.Parameters("@bh").Value = 0 Then
            MsgBox ("出错!,请重新打开再试一次!")
            Me.Visible = False
        End If
        
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = comQy.Text
        mod1.cmd.Parameters("@mt2") = ""
        mod1.cmd.Parameters("@mt3") = txtKhmc.Text & ":" & txtJe.Text
        mod1.cmd.Parameters("@mt5") = Me.Ywy
        mod1.cmd.Parameters("@mt6") = Me.Uid
        mod1.cmd.Parameters("@mlt1") = txtQM.Text '评审建议
        mod1.cmd.Parameters("@mm1").Value = Me.Lc
        mod1.cmd.Parameters("@mm2").Value = Fwid
        mod1.cmd.Parameters("@mm3") = 0
        '公司名称
        If companyId.Text = "上海豪曼制冷空调服务有限公司" Then
            mod1.cmd.Parameters("@mm5") = 1
        ElseIf companyId.Text = "上海鼎力制冷空调设备有限公司" Then
            mod1.cmd.Parameters("@mm5") = 2
        ElseIf companyId.Text = "上海杰升商贸有限公司" Then
            mod1.cmd.Parameters("@mm5") = 3
        End If

        If OptT1.Value = True Then
            mod1.cmd.Parameters("@mb1") = 1 '同意
        Else
            mod1.cmd.Parameters("@mb1") = 0 '拒绝
        End If
        mod1.cmd.Parameters("@md1") = Null
 
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

Private Sub cmdGx_Click()
On Error Resume Next
'If Jid = 0 Then Exit Sub
If txtXrq.Text = "" Then Exit Sub
If Len(Trim(txtZ.Text)) <> 5 Then Exit Sub
frmQm.Visible = False
        timZm = 7 '更新结算
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "财务到帐"
        mod1.cmd.Parameters("@NBLX") = "更新结算"
        mod1.cmd.Parameters("@bh") = lblAid.Caption
       
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = ""
        mod1.cmd.Parameters("@mt2") = ""
        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1").Value = Val(txtZ.Text)
        mod1.cmd.Parameters("@mm2").Value = Val(txtJJ.Text)
        mod1.cmd.Parameters("@mb1") = 0

        mod1.cmd.Parameters("@md1") = txtXrq.Text
 
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
If frmXX.Visible = True Then
    frmXX.Visible = False
End If

If LCUid = mod1.DHid Or mod1.DName = "马晓聪" Then
    cmdSave.Enabled = True

End If
If Lc = 1 And mod1.DName = LCRen Then
    txtKhmc.Locked = False
    dtpDate.Enabled = True
    txtJe.Locked = False
    comQy.Locked = False
    cmdDel.Enabled = True
ElseIf Lc = 2 Then
    txtBz.Locked = False
    frmXX.Visible = True
End If


End Sub

Private Sub cmdNQ_Click()

Dim oo As Integer

Dim ii As Integer


On Error Resume Next






If LCRen <> mod1.DName And Lc <> 100 Then
    MsgBox "此处应由" & LCRen & "签字! 请您不要再点"
    Exit Sub
End If

If cmdSave.Enabled = True Then
    MsgBox "请先将单子保存,再签上您的大名!"
    Exit Sub
End If

    frmQm.Visible = True
    cmdDing.Enabled = True
    
    If Me.Lc = 1 Then   '报销人只能签字，不能驳回。
        optT2.Enabled = False
        OptT1.Value = True
    Else
        optT2.Enabled = True
        OptT1.Value = False
        optT2.Value = False
    End If
If Lc = 100 And mod1.Bm = "商务部" Then '商务部在流程结束后，还能驳回修改
        optT2.Enabled = True
        OptT1.Value = False
        optT2.Value = True
End If
End Sub


Private Sub cmdSave_Click()

timZm = 1 '保存
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "财务到帐"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = lblAid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtKhmc.Text
    mod1.cmd.Parameters("@mt2") = comQy.Text
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mt4") = ""
    mod1.cmd.Parameters("@mt5") = ""

    mod1.cmd.Parameters("@mlt1") = txtBz.Text '备注
    mod1.cmd.Parameters("@mm1") = txtJe.Text
    '公司名称
    If companyId.Text = "上海豪曼制冷空调服务有限公司" Then
        mod1.cmd.Parameters("@mm2") = 1
    ElseIf companyId.Text = "上海鼎力制冷空调设备有限公司" Then
        mod1.cmd.Parameters("@mm2") = 2
    ElseIf companyId.Text = "上海杰升商贸有限公司" Then
        mod1.cmd.Parameters("@mm2") = 3
    End If
    mod1.cmd.Parameters("@mb1") = 0

    mod1.cmd.Parameters("@md1") = txtDate.Text

    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 2 Then '保存
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
End Sub

Private Sub dtgFk_Click()
dtgFKN.Row = dtgFk.Row
dtgFKN.Col = 3
Fid = Val(dtgFKN.Text)
End Sub

Private Sub dtgXL_Click()
dtgXLN.Row = dtgXL.Row
dtgXLN.Col = 3
Jid = Val(dtgXLN.Text)
dtgXLN.Col = 0: txtZ.Text = dtgXLN.Text
dtgXLN.Col = 2: txtJJ.Text = Val(dtgXLN.Text)
dtgXLN.Col = 1: txtXrq.Text = dtgXLN.Text
End Sub

Private Sub dtgXrq_CloseUp()
txtXrq.Text = dtgXrq.Value
End Sub


Private Sub dtpDate_CloseUp()
txtDate.Text = dtpDate.Value
End Sub


Private Sub Form_Click()
frmQm.Visible = False
frmXX.Visible = False
End Sub
Public Sub FKBound(Rb, Lb As Integer, HTZE As Single)
Dim FK As Single
Dim oo As Integer
Call FKFF
On Error Resume Next
For oo = 1 To Lb
    
    dtgFk.Row = oo
    dtgFk.Col = 0: dtgFk.Text = Rb(0, oo - 1): dtgFKN.Col = 0: dtgFKN.Text = Rb(0, oo - 1)
    dtgFk.Col = 2: dtgFk.Text = Rb(2, oo - 1): FK = Rb(2, oo - 1)
    dtgFk.Col = 1: dtgFk.Text = Str(Round(FK / HTZE, 2) * 100) & "%"
    dtgFk.Col = 3: dtgFk.Text = Rb(3, oo - 1)
    If Rb(4, oo - 1) = True Then
        dtgFk.Col = 4: dtgFk.Text = "是"
        dtgFk.Col = 0: dtgFk.Text = "款到发货"
        dtgFk.CellAlignment = 0
        dtgFKN.Col = 0: dtgFKN.Text = "款到发货"
    End If
    dtgFKN.Row = oo

    dtgFKN.Col = 1: dtgFKN.Text = Str(Round(FK / HTZE, 2) * 100) & "%"
    dtgFKN.Col = 2: dtgFKN.Text = Rb(2, oo - 1)
    dtgFKN.Col = 3: dtgFKN.Text = Rb(3, oo - 1)
    dtgFKN.Col = 4: dtgFKN.Text = dtgFk.Text

Next
End Sub
Public Sub FKFF()
dtgFk.Clear
dtgFKN.Clear
dtgFk.Rows = 30
dtgFk.Cols = 5
dtgFk.Row = 0
dtgFk.Col = 0: dtgFk.Text = "日期": dtgFk.CellFontBold = True
dtgFk.Col = 1: dtgFk.Text = "额度": dtgFk.CellFontBold = True
dtgFk.Col = 2: dtgFk.Text = "金额": dtgFk.CellFontBold = True
dtgFk.Col = 3: dtgFk.Text = "fid": dtgFk.CellFontBold = True
dtgFk.Col = 4: dtgFk.Text = "款到发货": dtgFk.CellFontBold = True

dtgFk.ColWidth(3) = 0
dtgFk.ColWidth(4) = 0
dtgFk.ColWidth(0) = 1100

dtgFKN.Rows = 30
dtgFKN.Cols = 5
End Sub
Private Sub Form_Load()
Me.Height = 6300
Me.Width = 10440
dtpDate.Value = mod1.DQda
dtgXrq.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Visible = False
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
ElseIf FMXCYBR.Visible = True Then
    Call FMXCYBR.REF(FMXCYBR.tt)
End If
End Sub


Private Sub frmCw_Click()
frmXX.Visible = False
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

If timZm = 1 Then '如果为添加合同评审
    cmdSave.Enabled = False

ElseIf timZm = 2 Then
    Me.Visible = False '删除
    If Dialog.Visible = True Then
        Dialog.Enabled = True
        Dialog.ZOrder 0
        Call mod1.refEnvent(1)
    End If
ElseIf timZm = 5 Then '签字
    Call QMBound(Val(lblAid.Caption))
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTX.Visible = True
    If Dialog.Visible = True Then
    Call mod1.refEnvent(1)
    End If
ElseIf timZm = 7 Or timZm = 8 Then
    tt = "select hid,rq,je,hxrq,jid from htADetail where aid=" & Val(lblAid.Caption)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    On Error Resume Next
    RC = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lc = UBound(RC, 2) + 1
    Call Me.XLBound(RC)
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
        cmdSave.Enabled = False
        If Val(lblAid.Caption) = 0 Then
            lblAid.Caption = Int(mod1.WP.Fields("mt1").Value)
        End If

    ElseIf timZm = 5 Then
        frmQm.Visible = False
        Me.Lc = mod1.WP.Fields("mm1").Value
        Fwid = mod1.WP.Fields("mm2").Value
        LCRen = mod1.WP.Fields("mt1").Value
        LCUid = mod1.WP.Fields("mt2").Value
        lblTX.Caption = "当前流程至:" & LCRen
        If Me.Lc = 100 Then
            lblTX.Caption = "流程结束"
        End If
        
        Call QMBound(Val(lblZid.ToolTipText))
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



Public Sub Bound(Aid As Long)
Dim Ra
Dim RC

Dim Cid As Integer
Call Qing
tt = "select khmc,dzrq,je,bz,lc,lcren,lcuid,ywy,uid,fwid,qy,companyid from htAcount where aid=" & Aid & ";" & _
    "select htbh,rq,je,jid from htADView where aid=" & Aid & " order by rq desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Lc = UBound(RC, 2) + 1
txtKhmc.Text = Ra(0, 0)
txtDate.Text = Ra(1, 0)
txtJe.Text = Ra(2, 0)
txtBz.Text = Ra(3, 0)
Lc = Ra(4, 0)
LCRen = Ra(5, 0)
LCUid = Ra(6, 0)
Ywy = Ra(7, 0)
Uid = Ra(8, 0)
Fwid = Ra(9, 0)
comQy.Text = Ra(10, 0)
Cid = Ra(11, 0)
If Cid = 1 Then
    companyId.Text = "上海豪曼制冷空调服务有限公司"
ElseIf Cid = 2 Then
    companyId.Text = "上海鼎力制冷空调设备有限公司"
ElseIf Cid = 3 Then
    companyId.Text = "上海杰升商贸有限公司"
End If
lblAid.Caption = Aid
lblTX.Caption = "目前流程至:" & LCRen
If Lc = 100 Then
    lblTX.Caption = "流程结束!"
End If
Call Me.XLBound(RC)
Call QMBound(Aid)
End Sub

Public Sub dtgXlFF()
dtgXL.Clear
dtgXL.Cols = 4
dtgXL.Rows = 30
dtgXL.Row = 0
dtgXL.Col = 0: dtgXL.Text = "合同编号": dtgXL.CellFontBold = True
dtgXL.Col = 1: dtgXL.Text = "结算日期": dtgXL.CellFontBold = True
dtgXL.Col = 2: dtgXL.Text = "结算金额": dtgXL.CellFontBold = True


dtgXLN.Clear
dtgXLN.Cols = 4
dtgXLN.Rows = 30
dtgXLN.Row = 0
dtgXLN.Col = 0: dtgXLN.Text = "合同编号": dtgXLN.CellFontBold = True
dtgXLN.Col = 1: dtgXLN.Text = "结算日期": dtgXLN.CellFontBold = True
dtgXLN.Col = 2: dtgXLN.Text = "结算金额": dtgXLN.CellFontBold = True

dtgXL.ColWidth(3) = 0
dtgXL.ColWidth(0) = 1900
End Sub


Public Sub XLBound(RC)
On Error Resume Next
Dim oo As Integer
Dim Xg As Double
Dim Lc As Integer
Lc = UBound(RC, 2) + 1
For oo = 1 To Lc
    dtgXL.Row = oo
    dtgXL.Col = 0: dtgXL.Text = RC(0, oo - 1)
    dtgXL.Col = 1: dtgXL.Text = RC(1, oo - 1)
    dtgXL.Col = 2: dtgXL.Text = RC(2, oo - 1)
    Xg = Xg + RC(2, oo - 1)
    dtgXL.Col = 3: dtgXL.Text = RC(3, oo - 1)
'''    dtgXL.Col = 4: dtgXL.Text = RC(4, oo - 1)
'''    dtgXL.Col = 5: dtgXL.Text = RC(5, oo - 1)
    
    dtgXLN.Row = oo
    dtgXLN.Col = 0: dtgXLN.Text = RC(0, oo - 1)
    dtgXLN.Col = 1: dtgXLN.Text = RC(1, oo - 1)
    dtgXLN.Col = 2: dtgXLN.Text = RC(2, oo - 1)
    dtgXLN.Col = 3: dtgXLN.Text = RC(3, oo - 1)
'''    dtgXLN.Col = 4: dtgXLN.Text = RC(4, oo - 1)
'''    dtgXLN.Col = 5: dtgXLN.Text = RC(5, oo - 1)
Next
dtgXL.Row = oo
dtgXL.Col = 1: dtgXL.Text = "小计"
dtgXL.Col = 2: dtgXL.Text = Xg
If Xg > Me.txtJe.Text Then
    dtgXL.CellForeColor = &HFF&
End If
End Sub
