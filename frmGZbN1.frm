VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGZbN1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "SERVICE SALES WEEKLY REPORT"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   8940
      TabIndex        =   45
      Top             =   6300
      Visible         =   0   'False
      Width           =   6315
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   48
         Top             =   870
         Width           =   675
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   47
         Top             =   510
         Width           =   705
      End
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1305
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Top             =   300
         Width           =   4965
      End
   End
   Begin VB.CommandButton cmdRight 
      Height          =   495
      Left            =   12090
      Picture         =   "frmGZbN1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8700
      Width           =   495
   End
   Begin VB.CommandButton cmdLeft 
      Height          =   495
      Left            =   11580
      Picture         =   "frmGZbN1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8700
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   405
      Left            =   12180
      TabIndex        =   42
      Top             =   150
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   714
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4860
      Top             =   2640
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5670
      Top             =   2610
   End
   Begin VB.TextBox txtED 
      BackColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   7110
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "frmGZbN1.frx":0884
      Top             =   7290
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton cmdCreate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "新建"
      Height          =   555
      Left            =   12660
      Picture         =   "frmGZbN1.frx":088A
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8700
      Width           =   675
   End
   Begin VB.TextBox txtBF 
      BackColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   14040
      TabIndex        =   33
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "审核"
      Height          =   525
      Left            =   10230
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8700
      Width           =   855
   End
   Begin VB.TextBox txtBxid 
      BackColor       =   &H00C0FFC0&
      Height          =   270
      Left            =   14040
      TabIndex        =   29
      Top             =   420
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgD 
      Height          =   1605
      Left            =   0
      TabIndex        =   27
      Top             =   7620
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   2831
      _Version        =   393216
      BackColor       =   12648447
      Rows            =   10
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648447
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "保存"
      Height          =   555
      Left            =   13950
      Picture         =   "frmGZbN1.frx":0CCC
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "保存"
      Top             =   8700
      Width           =   615
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H00C0FFC0&
      Caption         =   "修改"
      Height          =   555
      Left            =   13350
      Picture         =   "frmGZbN1.frx":1336
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "修改"
      Top             =   8700
      Width           =   585
   End
   Begin VB.TextBox txtBz6 
      BackColor       =   &H00FFFFC0&
      Height          =   1455
      Left            =   10200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   4650
      Width           =   5025
   End
   Begin VB.TextBox txtBz5 
      BackColor       =   &H00FFFFC0&
      Height          =   1125
      Left            =   10200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   2970
      Width           =   5025
   End
   Begin VB.TextBox txtBz4 
      BackColor       =   &H00FFFFC0&
      Height          =   1395
      Left            =   10200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   1110
      Width           =   5025
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgA 
      Height          =   1815
      Left            =   0
      TabIndex        =   9
      Top             =   780
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3201
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
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00C0FFFF&
      Caption         =   "返回"
      Height          =   555
      Left            =   14580
      Picture         =   "frmGZbN1.frx":1640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8700
      Width           =   645
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgB 
      Height          =   1905
      Left            =   0
      TabIndex        =   17
      Top             =   2970
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3360
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
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgC 
      Height          =   1995
      Left            =   0
      TabIndex        =   19
      Top             =   5250
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3519
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
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   2115
      Left            =   10200
      TabIndex        =   30
      Top             =   6510
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   3731
      _Version        =   393216
      BackColor       =   12640511
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
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label lblTX 
      BackStyle       =   0  'Transparent
      Caption         =   "流程至:"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   11250
      TabIndex        =   50
      Top             =   6270
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Label lblLR 
      Caption         =   "lblLR"
      Height          =   195
      Left            =   6930
      TabIndex        =   41
      Top             =   5070
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lblFR 
      Caption         =   "lblFR"
      Height          =   195
      Left            =   4500
      TabIndex        =   40
      Top             =   5040
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblRen 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   2430
      TabIndex        =   39
      Top             =   150
      Width           =   1515
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      Height          =   225
      Left            =   1890
      TabIndex        =   38
      Top             =   150
      Width           =   585
   End
   Begin VB.Label lblGid 
      Caption         =   "lblGid"
      Height          =   195
      Left            =   5640
      TabIndex        =   37
      Top             =   7410
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblHK 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   11520
      TabIndex        =   36
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "回款天数："
      Height          =   165
      Left            =   10500
      TabIndex        =   35
      Top             =   480
      Width           =   945
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "报销金额"
      Height          =   195
      Left            =   13140
      TabIndex        =   32
      Top             =   150
      Width           =   795
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "审阅栏"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10350
      TabIndex        =   28
      Top             =   6270
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "单据编号"
      Height          =   195
      Left            =   13140
      TabIndex        =   26
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblT 
      BackStyle       =   0  'Transparent
      Caption         =   "    在编辑中如需换行,可以按CTRL+回车键,按回车键或双击文本框,则退出编辑状态"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3360
      TabIndex        =   25
      Top             =   540
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "7/ ACTION PLAN NEXT WEEK 下周行动计划"
      Height          =   225
      Left            =   90
      TabIndex        =   24
      Top             =   7350
      Width           =   4065
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "3/Maintenance Platform维护平台（20%）"
      Height          =   225
      Left            =   90
      TabIndex        =   20
      Top             =   4980
      Width           =   3735
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "2/Working Platform工作平台（50%）"
      Height          =   225
      Left            =   90
      TabIndex        =   18
      Top             =   2730
      Width           =   3735
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "1/Marketing Platform市场平台（30%）"
      Height          =   225
      Left            =   60
      TabIndex        =   16
      Top             =   510
      Width           =   3735
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "6/ Service Quality Feedback 售后服务质量问题"
      Height          =   165
      Left            =   10320
      TabIndex        =   14
      Top             =   4320
      Width           =   4065
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "5/ MARKET CHANGES / COMPETITOR ACTIVETIES 维修销售市场/竞争对手资料"
      Height          =   375
      Left            =   10320
      TabIndex        =   12
      Top             =   2580
      Width           =   3945
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "4 CRITICAL ISSUES  运作问题和困难"
      Height          =   225
      Left            =   10320
      TabIndex        =   10
      Top             =   840
      Width           =   3435
   End
   Begin VB.Label lblYZ 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   11520
      TabIndex        =   8
      Top             =   120
      Width           =   705
   End
   Begin VB.Label lblWZB 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   9210
      TabIndex        =   7
      Top             =   150
      Width           =   855
   End
   Begin VB.Label lblGzB 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   6420
      TabIndex        =   6
      Top             =   150
      Width           =   705
   End
   Begin VB.Label lblBm 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   570
      TabIndex        =   5
      Top             =   150
      Width           =   885
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "应收帐款："
      Height          =   225
      Left            =   10500
      TabIndex        =   4
      Top             =   150
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "已完成指标（开票）："
      Height          =   345
      Left            =   7320
      TabIndex        =   3
      Top             =   150
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "个人销售指标"
      Height          =   225
      Left            =   5160
      TabIndex        =   2
      Top             =   150
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "部门"
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   150
      Width           =   435
   End
End
Attribute VB_Name = "frmGZbN1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer '(1保存)
Dim ED As Integer '表格编辑框所处的位置
Dim LCRen As String
Dim LCUid As String
Public Lc As Integer
Dim Fwid As Long
Public Sub QMBound(Gid As Long)
Dim Ra: Dim La
Dim ii As Integer: Dim oo As Integer
Dim tt As String
On Error Resume Next

tt = "select trq,ywy,zn,bz,tf from pizu where bh='" & Gid & "' and yid=40 order by pid desc"
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



End Sub
Public Sub GetWeek(mtA As Date)
Select Case DatePart("w", mtA)
Case 1 '星期日
lblFR.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) - 6)
lblLR.Caption = mtA
Case 2 '星期一
lblFR.Caption = mtA
lblLR.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) + 6)
Case 3
lblFR.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) - 1)
lblLR.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) + 5)
Case 4
lblFR.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) - 2)
lblLR.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) + 4)
Case 5
lblFR.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) - 3)
lblLR.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) + 3)
Case 6
lblFR.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) - 4)
lblLR.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) + 2)
Case 7
lblFR.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) - 5)
lblLR.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) + 1)
End Select
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
 dtgP.ColWidth(3) = 2115: dtgP.ColWidth(4) = 525
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub
Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
frmZu.ZOrder 0
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub



Private Sub cmdCreate_Click()
Dim tt As String
Dim Ra
Dim Frq As Date
Dim LRQ As Date
If mod1.DHid <> lblRen.ToolTipText Or Lc > 1 Or Val(lblGid.Caption) > 0 Then
    Exit Sub
End If
Call Qing
Call Qing
On Error GoTo frmgzbn1ERR
''''''获取日期
'''''tt = "select top 1 fr,lr from SalesReport where uid='" & mod1.DHid & "' order by gid desc"
'''''Set mod1.HTP = CreateObject("adodb.recordset")
'''''mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'''''If mod1.HTP.BOF = True Then
'''''    Call GetWeek(mod1.DQda)
'''''Else
'''''    Ra = mod1.HTP.GetRows
'''''    mod1.HTP.Close
'''''    Frq = Ra(0, 0)
'''''    LRQ = Ra(1, 0)
'''''    lblFR.Caption = DateSerial(Year(Frq), Month(Frq), Day(Frq) + 7)
'''''    lblLR.Caption = DateSerial(Year(LRQ), Month(LRQ), Day(LRQ) + 7)
'''''End If
'''''Set mod1.HTP = Nothing
Me.Caption = "SERVICE SALES WEEKLY REPORT     " & lblFR.Caption & "  to   " & lblLR.Caption
dtgA.SelectionMode = flexSelectionFree
dtgB.SelectionMode = flexSelectionFree
dtgC.SelectionMode = flexSelectionFree
dtgD.SelectionMode = flexSelectionFree
dtgA.Enabled = True
dtgB.Enabled = True
dtgC.Enabled = True
dtgD.Enabled = True
txtBz4.Locked = False
txtBz5.Locked = False
txtBz6.Locked = False
txtBF.Locked = False
txtBxid.Locked = False
lblT.Visible = True
lblT.Caption = "    双击单元格可进行编辑"
lblBM.Caption = mod1.Bm: lblBM.ToolTipText = mod1.Bmid
lblRen.Caption = mod1.DName: lblRen.ToolTipText = mod1.DHid
cmdSave.Enabled = True
Exit Sub
frmgzbn1ERR:
MsgBox "网络故障，请再试一次！"
End Sub

Private Sub cmdDing_Click()
Dim tt As String
On Error Resume Next
If Lc = 0 Then
    Exit Sub
End If
If optT2.Value = True And txtQM.Text = "" Then
    MsgBox ("请您一定要告诉拒绝我的理由!  :) ")
    Exit Sub
End If
frmQm.Visible = False


        timZm = 2 '签字
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "工作报告"
        mod1.cmd.Parameters("@NBLX") = "签字"
        mod1.cmd.Parameters("@bh") = lblGid.Caption
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = lblRen.Caption
        mod1.cmd.Parameters("@mt2") = lblRen.ToolTipText
        mod1.cmd.Parameters("@mt3") = lblBM.Caption
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
        mod1.cmd.Parameters("@mm1").Value = Me.Lc
        mod1.cmd.Parameters("@mm2").Value = Fwid
        mod1.cmd.Parameters("@mm3") = 0
        mod1.cmd.Parameters("@mm4") = 0
        mod1.cmd.Parameters("@mm5") = 0
        mod1.cmd.Parameters("@mm6") = 0
        mod1.cmd.Parameters("@mm7") = 0
        mod1.cmd.Parameters("@mm8") = 0
        mod1.cmd.Parameters("@mm9") = 0
        mod1.cmd.Parameters("@mm10").Value = 0
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

Private Sub cmdLeft_Click()
Dim Frq As Date
Dim Uid As String
Uid = lblRen.ToolTipText
Call frmGZbN1.Qing
Call frmGZbN1.Qing
Frq = DateSerial(Year(lblFR.Caption), Month(lblFR.Caption), Day(lblFR.Caption) - 7)
Call frmGZbN1.Bound(Uid, Frq)
End Sub


Private Sub cmdMod_Click()
dtgA.Enabled = True
dtgB.Enabled = True
dtgC.Enabled = True
dtgD.Enabled = True
If mod1.DHid <> lblRen.ToolTipText Or Lc > 1 Then
    Exit Sub
End If
dtgA.SelectionMode = flexSelectionFree
dtgB.SelectionMode = flexSelectionFree
dtgC.SelectionMode = flexSelectionFree
dtgD.SelectionMode = flexSelectionFree
txtBF.Locked = False
txtBxid.Locked = False
txtBz4.Locked = False
txtBz5.Locked = False
txtBz6.Locked = False
lblT.Visible = True
lblT.Caption = "    双击单元格可进行编辑"
cmdSave.Enabled = True

End Sub

Private Sub cmdNQ_Click()

Dim tt As String
Dim oo As Integer

Dim ii As Integer


On Error Resume Next






If LCRen <> mod1.DName Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字! 请您不要再点"
    Exit Sub
End If
If Lc = 100 Then

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
    Exit Sub















End Sub

Private Sub cmdRight_Click()
Dim Frq As Date
Dim Uid As String
Uid = lblRen.ToolTipText
Call frmGZbN1.Qing
Call frmGZbN1.Qing
Frq = DateSerial(Year(lblFR.Caption), Month(lblFR.Caption), Day(lblFR.Caption) + 7)
Call frmGZbN1.Bound(Uid, Frq)
End Sub

Private Sub cmdSave_Click()
Dim oo As Integer: Dim ii As Integer
Dim tt As String
Dim arrayDTG(20, 100) As String
If txtEd.Visible = True Then Exit Sub '编辑一半的状态,不能保存
dtgA.SelectionMode = flexSelectionByRow
dtgB.SelectionMode = flexSelectionByRow
dtgC.SelectionMode = flexSelectionByRow
dtgD.SelectionMode = flexSelectionByRow
txtBF.Locked = True
txtBxid.Locked = True
txtBz4.Locked = True
txtBz5.Locked = True
txtBz6.Locked = True
Call GDui
cmdSave.Enabled = False
lblT.Visible = False
For oo = 0 To dtgA.Rows - 3
    dtgA.Row = oo + 1
    For ii = 0 To dtgA.Cols - 1
        dtgA.Col = ii

        arrayDTG(ii + 1, oo) = dtgA.Text
        If ii = 0 Then
            arrayDTG(ii, oo) = "A"
        End If
    Next
Next
For oo = 0 To dtgB.Rows - 3
    dtgB.Row = oo + 1
    For ii = 0 To dtgB.Cols - 1
        dtgB.Col = ii
        arrayDTG(ii + 1, oo + Val(dtgA.ToolTipText) - 2) = dtgB.Text
        If ii = 0 Then
            arrayDTG(ii, oo + Val(dtgA.ToolTipText) - 2) = "B"
        End If
    Next
Next
For oo = 0 To dtgC.Rows - 3
    dtgC.Row = oo + 1
    For ii = 0 To dtgC.Cols - 1
        dtgC.Col = ii
        arrayDTG(ii + 1, oo + Val(dtgA.ToolTipText) + Val(dtgB.ToolTipText) - 2 - 2) = dtgC.Text
        If ii = 0 Then
            arrayDTG(ii, oo + Val(dtgA.ToolTipText) + Val(dtgB.ToolTipText) - 2 - 2) = "C"
        End If
    Next
Next
For oo = 0 To dtgD.Rows - 3
    dtgD.Row = oo + 1
    For ii = 0 To dtgD.Cols - 1
        dtgD.Col = ii
        arrayDTG(ii + 1, oo + Val(dtgA.ToolTipText) + Val(dtgB.ToolTipText) + Val(dtgC.ToolTipText) - 6) = dtgD.Text
        If ii = 0 Then
            arrayDTG(ii, oo + Val(dtgA.ToolTipText) + Val(dtgB.ToolTipText) + Val(dtgC.ToolTipText) - 6) = "D"
        End If
    Next
Next



timZm = 1 '保存合同
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAddArray"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "工作报告"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = lblGid.Caption
    mod1.cmd.Parameters("@ywy") = lblRen.Caption
    mod1.cmd.Parameters("@uid") = lblRen.ToolTipText
    mod1.cmd.Parameters("@mt1") = txtBxid.Text '报销单编号
    mod1.cmd.Parameters("@mt2") = lblBM.ToolTipText
    mod1.cmd.Parameters("@mt25") = ""
    mod1.cmd.Parameters("@mlt1") = txtBz4.Text
    mod1.cmd.Parameters("@mlt2") = txtBz5.Text
    mod1.cmd.Parameters("@mlt3") = txtBz6.Text
    mod1.cmd.Parameters("@mlt4") = ""
    mod1.cmd.Parameters("@mlt5") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtBF.Text) '报销单金额
    mod1.cmd.Parameters("@mm20") = 0
    mod1.cmd.Parameters("@mb4") = 0
    mod1.cmd.Parameters("@mb5") = 0
    mod1.cmd.Parameters("@md1") = lblFR.Caption
    mod1.cmd.Parameters("@md2") = lblLR.Caption
    mod1.cmd.Parameters("@md3") = Null
    mod1.cmd.Parameters("@md4") = Null
    mod1.cmd.Parameters("@md5") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
frmGZBNERR1:
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"

        Exit Sub
    Else '提交成功,等待系统中心处理数据
        For oo = 0 To (Val(dtgA.ToolTipText) + Val(dtgB.ToolTipText) + Val(dtgC.ToolTipText) + Val(dtgC.ToolTipText) - 9)
            tt = tt & "insert into hmtext.dbo.MLarray (A00,A01,A02,A03,A04,A05,A06,A07,zid) values ('" & arrayDTG(0, oo) & "','" & arrayDTG(1, oo) & "','" & arrayDTG(2, oo) & _
                    "','" & arrayDTG(3, oo) & "','" & arrayDTG(4, oo) & "','" & arrayDTG(5, oo) & "','" & arrayDTG(6, oo) & "','" & arrayDTG(7, oo) & "'," & mod1.Zid & ")"

                tt = tt & ";"

        Next
        tt = tt & "update hmtext.dbo.ML set cf=0,arrayF=1 where zid=" & mod1.Zid
        Set mod1.HTP = CreateObject("adodb.recordset")
        On Error GoTo frmGZBNERR1
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        'mod1.HTP.Close
        Set mod1.HTP = Nothing
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
        'MsgBox "OK!"
    End If


Set mod1.cmd = Nothing


End Sub


Private Sub dtgA_Click()
frmQm.Visible = False
End Sub

Private Sub dtgA_DblClick()
If dtgA.SelectionMode = flexSelectionByRow Then Exit Sub
If dtgA.Row > Val(dtgA.ToolTipText) Then Exit Sub
txtEd.Left = dtgA.CellLeft + dtgA.Left
txtEd.Top = dtgA.CellTop + dtgA.Top
txtEd.Width = dtgA.CellWidth
txtEd.Height = dtgA.CellHeight
txtEd.Visible = True
txtEd.Text = dtgA.Text
txtEd.SetFocus
dtgA.Enabled = False
lblT.Visible = True
lblT.Caption = "    在编辑中如需换行,可以按CTRL+回车键,按回车键或双击文本框,则退出编辑状态"
ED = 1
txtEd.BackColor = dtgA.CellBackColor
End Sub



Private Sub dtgB_Click()
frmQm.Visible = False
End Sub

Private Sub dtgB_DblClick()

If dtgB.SelectionMode = flexSelectionByRow Then Exit Sub
If dtgB.Row > Val(dtgB.ToolTipText) Then Exit Sub
txtEd.Left = dtgB.CellLeft + dtgB.Left
txtEd.Top = dtgB.CellTop + dtgB.Top
txtEd.Width = dtgB.CellWidth
txtEd.Height = dtgB.CellHeight
txtEd.Visible = True
txtEd.Text = dtgB.Text
txtEd.SetFocus
dtgB.Enabled = False
lblT.Visible = True
lblT.Caption = "    在编辑中如需换行,可以按CTRL+回车键,按回车键或双击文本框,则退出编辑状态"
ED = 2
txtEd.BackColor = dtgB.CellBackColor
End Sub

Private Sub dtgC_Click()
frmQm.Visible = False
End Sub

Private Sub dtgC_DblClick()
If dtgC.SelectionMode = flexSelectionByRow Then Exit Sub
If dtgC.Row > Val(dtgC.ToolTipText) Then Exit Sub
txtEd.Left = dtgC.CellLeft + dtgC.Left
txtEd.Top = dtgC.CellTop + dtgC.Top
txtEd.Width = dtgC.CellWidth
txtEd.Height = dtgC.CellHeight
txtEd.Visible = True
txtEd.Text = dtgC.Text
txtEd.SetFocus
dtgC.Enabled = False
lblT.Visible = True
lblT.Caption = "    在编辑中如需换行,可以按CTRL+回车键,按回车键或双击文本框,则退出编辑状态"
ED = 3
txtEd.BackColor = dtgC.CellBackColor
End Sub


Private Sub dtgD_Click()
frmQm.Visible = False
End Sub

Private Sub dtgD_DblClick()
If dtgD.SelectionMode = flexSelectionByRow Then Exit Sub
If dtgD.Row > Val(dtgD.ToolTipText) Then Exit Sub
txtEd.Left = dtgD.CellLeft + dtgD.Left
txtEd.Top = dtgD.CellTop + dtgD.Top
txtEd.Width = dtgD.CellWidth
txtEd.Height = dtgD.CellHeight
txtEd.Visible = True
txtEd.Text = dtgD.Text
txtEd.SetFocus
dtgD.Enabled = False
lblT.Visible = True
lblT.Caption = "    在编辑中如需换行,可以按CTRL+回车键,按回车键或双击文本框,则退出编辑状态"
ED = 4
txtEd.BackColor = dtgD.CellBackColor
End Sub


Private Sub dtgP_Click()
frmQm.Visible = False
End Sub

Private Sub Form_Click()
frmQm.Visible = False
End Sub

Private Sub Form_Load()
Me.Left = 0: Me.Top = 0
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Dim ii As Integer: Dim oo As Integer
'''''dtgNa.Left = dtgA.Left: dtgNa.Top = dtgA.Top
'''''dtgNa.Rows = dtgA.Rows: dtgNa.Cols = dtgA.Cols
dtgA.RowHeight(0) = dtgA.RowHeight(0) * 2
dtgB.RowHeight(0) = dtgB.RowHeight(0) * 2.5
dtgC.RowHeight(0) = dtgC.RowHeight(0) * 2
Call dtgAFF
Call dtgBFF
Call dtgCFF
Call dtgDFF
Call dtgPFF
dtgA.ColWidth(7) = 0
dtgB.ColWidth(7) = 0
dtgC.ColWidth(7) = 0
dtgD.ColWidth(5) = 0
frmQm.Top = 7460
frmQm.Left = 8940
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
frmZu.Enabled = True
frmZu.ZOrder 0
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
Cancel = True
End Sub


Private Sub timQuit_Timer()
Dim oo As Integer
Dim ii As Integer
On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
Dim tt As String
If timZm = 1 Then '如果为添加合同评审
ElseIf timZm = 2 Then '签字
    cmdDing.Enabled = True
    txtQM.Text = ""
    frmQm.Visible = False
    lblTx.Visible = True
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

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2,mt3 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdSave.Enabled = False
    Else: timZm = 2
        frmQm.Visible = False
        Me.Lc = mod1.WP.Fields("mm1").Value
        Fwid = mod1.WP.Fields("mm2").Value
        LCRen = mod1.WP.Fields("mt1").Value
        LCUid = mod1.WP.Fields("mt2").Value
        LZw = mod1.WP.Fields("mt3").Value
            lblTx.Caption = "流程至" & LZw & ": " & LCRen
       Call QMBound(Val(lblGid.Caption))
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

Private Sub txtBz4_Click()
frmQm.Visible = False
End Sub


Private Sub txtBz5_Click()
frmQm.Visible = False
End Sub


Private Sub txtBz6_Click()
frmQm.Visible = False
End Sub


Private Sub txtEd_DblClick()
Dim DH As Single
Dim Xmmc As String: Dim PF As Integer
Dim oo As Integer
    txtEd.Visible = False
    Select Case ED
    Case 1
        dtgA.Text = txtEd.Text
        dtgA.Enabled = True
        DH = 255 * mod1.HH(txtEd.Text, UpInt(dtgA.CellWidth / 200))
        If DH > dtgA.RowHeight(dtgA.Row) Then
            dtgA.RowHeight(dtgA.Row) = DH
        End If
        lblT.Caption = "    双击单元格进行编辑"
        oo = dtgA.Col
        dtgA.Col = 0
        Xmmc = Trim(dtgA.Text)
        dtgA.Col = 5
        PF = Val(dtgA.Text)
        dtgA.Col = oo
        If Xmmc = "" Or PF = 0 Or dtgA.Row + 1 <> Val(dtgA.ToolTipText) Then Exit Sub
        dtgA.Rows = dtgA.Rows + 1
        dtgA.ToolTipText = dtgA.Rows
    Case 2
        dtgB.Text = txtEd.Text
        dtgB.Enabled = True
        DH = 255 * mod1.HH(txtEd.Text, UpInt(dtgB.CellWidth / 200))
        If DH > dtgB.RowHeight(dtgB.Row) Then
            dtgB.RowHeight(dtgB.Row) = DH
        End If
        lblT.Caption = "    双击单元格进行编辑"
        oo = dtgB.Col
        dtgB.Col = 0
        Xmmc = Trim(dtgB.Text)
        dtgB.Col = 5
        PF = Val(dtgB.Text)
        dtgB.Col = oo
        If Xmmc = "" Or PF = 0 Or dtgB.Row + 1 <> Val(dtgB.ToolTipText) Then Exit Sub
        dtgB.Rows = dtgB.Rows + 1
        dtgB.ToolTipText = dtgB.Rows
    Case 3
        dtgC.Text = txtEd.Text
        dtgC.Enabled = True
        DH = 255 * mod1.HH(txtEd.Text, UpInt(dtgC.CellWidth / 200))
        If DH > dtgC.RowHeight(dtgC.Row) Then
            dtgC.RowHeight(dtgC.Row) = DH
        End If
        lblT.Caption = "    双击单元格进行编辑"
        oo = dtgC.Col
        dtgC.Col = 0
        Xmmc = Trim(dtgC.Text)
        dtgC.Col = 5
        PF = Val(dtgC.Text)
        dtgC.Col = oo
        If Xmmc = "" Or PF = 0 Or dtgC.Row + 1 <> Val(dtgC.ToolTipText) Then Exit Sub
        dtgC.Rows = dtgC.Rows + 1
        dtgC.ToolTipText = dtgC.Rows
    Case 4
        dtgD.Text = txtEd.Text
        dtgD.Enabled = True
        DH = 255 * mod1.HH(txtEd.Text, UpInt(dtgD.CellWidth / 200))
        If DH > dtgD.RowHeight(dtgD.Row) Then
            dtgD.RowHeight(dtgD.Row) = DH
        End If
        lblT.Caption = "    双击单元格进行编辑"
        oo = dtgD.Col
        dtgD.Col = 0
        Xmmc = Trim(dtgD.Text)
''''''        dtgD.Col = 5
''''''        PF = Val(dtgD.Text)
        dtgD.Col = oo
        If Xmmc = "" Or dtgD.Row + 1 <> Val(dtgD.ToolTipText) Then Exit Sub
        dtgD.Rows = dtgD.Rows + 1
        dtgD.ToolTipText = dtgD.Rows
    End Select
End Sub

Private Sub txtED_KeyDown(KeyCode As Integer, Shift As Integer)
Dim DH As Single
Dim Xmmc As String: Dim PF As Integer
Dim oo As Integer
If KeyCode = 13 And Shift = 0 Then
'''''    dtgA.Text = txtED.Text
'''''    txtED.Visible = False
'''''    dtgA.Enabled = True
'''''    DH = 255 * mod1.HH(txtED.Text, UpInt(dtgA.CellWidth / 270))
'''''    If DH > dtgA.RowHeight(dtgA.Row) Then
'''''        dtgA.RowHeight(dtgA.Row) = DH
'''''    End If
'''''    lblT.Caption = "    双击单元格进行编辑"
'''''    oo = dtgA.Col
'''''    dtgA.Col = 0
'''''    Xmmc = Trim(dtgA.Text)
'''''    dtgA.Col = 5
'''''    PF = Val(dtgA.Text)
'''''    dtgA.Col = oo
'''''    If Xmmc = "" Or PF = 0 Or dtgA.Row + 1 <> Val(dtgA.ToolTipText) Then Exit Sub
'''''    dtgA.Rows = dtgA.Rows + 1
'''''    dtgA.ToolTipText = dtgA.Rows
    Call txtEd_DblClick
End If
'''''If KeyCode = 13 And Shift = 2 Then
'''''    txtEd.Height = txtEd.Height + 150
'''''End If

End Sub



Public Sub Qing()

lblGzB.Caption = "" '个人指标
lblWZB.Caption = "" '完成指标
lblYZ.Caption = "" '应收帐款
lblHK.Caption = "" '回款天数
txtBxid.Text = ""
txtBF.Text = ""
txtBz4.Text = ""
txtBz5.Text = ""
txtBz6.Text = ""
Call dtgAFF
Call dtgBFF
Call dtgCFF
Call dtgDFF
Call dtgPFF
frmGZbN1.dtgA.SelectionMode = flexSelectionByRow
frmGZbN1.dtgB.SelectionMode = flexSelectionByRow
frmGZbN1.dtgC.SelectionMode = flexSelectionByRow
frmGZbN1.dtgD.SelectionMode = flexSelectionByRow

dtgA.ToolTipText = 2 '数据行数
dtgB.ToolTipText = 2 '数据行数
dtgC.ToolTipText = 2 '数据行数
dtgD.ToolTipText = 2 '数据行数
dtgA.Rows = 2
dtgB.Rows = 2
dtgC.Rows = 2
dtgD.Rows = 2
lblGid.Caption = ""
lblBM.Caption = ""
lblBM.ToolTipText = ""
lblRen.ToolTipText = ""
lblRen.Caption = ""
txtBF.Locked = True
txtBxid.Locked = True
cmdSave.Enabled = False
txtEd.Visible = False
txtBz4.Locked = True
txtBz5.Locked = True
txtBz6.Locked = True
lblTx.Visible = False
Lc = 0
LCRen = ""
LCUid = ""
Fwid = 0
txtQM.Text = ""
OptT1.Value = False
optT2.Value = False
OptT1.Enabled = True
optT2.Enabled = True
lblT.Visible = False
End Sub

Public Sub dtgAFF()
Dim oo As Integer
dtgA.Clear
For oo = 1 To dtgA.Rows - 1
    dtgA.RowHeight(oo) = dtgP.RowHeight(0)
Next
dtgA.Row = 0
dtgA.Col = 5: dtgA.Text = "评分" & Chr(13) & Chr(10) & "总分36"
dtgA.Col = 0: dtgA.Text = "客户名称" & Chr(13) & Chr(10) & "Customer Name"
dtgA.Col = 1: dtgA.Text = "金额" & Chr(13) & Chr(10) & "Amount"
dtgA.Col = 2: dtgA.Text = "竞争对手及报价"
dtgA.Col = 3: dtgA.Text = "机组情况"
dtgA.Col = 4: dtgA.Text = "客户意向"
dtgA.Col = 5: dtgA.Text = "评分" & Chr(13) & Chr(10) & "总分36"

dtgA.ColWidth(0) = 1900 ': dtgNa.ColWidth(0) = 1900
dtgA.ColWidth(1) = 770 ': dtgNa.ColWidth(1) = 770
dtgA.ColWidth(2) = 1665 ': dtgNa.ColWidth(2) = 1665
dtgA.ColWidth(3) = 1905 ': dtgNa.ColWidth(3) = 1905
dtgA.ColWidth(4) = 2505 ': dtgNa.ColWidth(4) = 2505
dtgA.ColWidth(6) = 0

'''''''For oo = 0 To dtgA.Rows - 1
'''''''    dtgA.Row = oo: dtgNa.Row = oo
'''''''    For ii = 0 To 5
'''''''        dtgA.Col = ii: dtgNa.Col = ii
'''''''        dtgNa.Text = dtgA.Text
''''''''''        dtgNa.CellHeight = dtgA.CellHeight
''''''''''        dtgNa.CellWidth = dtgA.CellWidth
'''''''    Next
'''''''Next



End Sub

Public Sub dtgBFF()
Dim oo As Integer
dtgB.Clear
For oo = 1 To dtgB.Rows - 1
    dtgB.RowHeight(oo) = dtgP.RowHeight(0)
Next
dtgB.Row = 0
dtgB.Col = 5: dtgB.Text = "评分" & Chr(13) & Chr(10) & "总分100"
dtgB.Col = 0: dtgB.Text = "客户名称" & Chr(13) & Chr(10) & "Customer Name"
dtgB.Col = 1: dtgB.Text = "金额" & Chr(13) & Chr(10) & "Amount" & Chr(13) & Chr(10) & "和MF"
dtgB.Col = 2: dtgB.Text = "竞争对手及报价"
dtgB.Col = 3: dtgB.Text = "需要公司何种支持"
dtgB.Col = 4: dtgB.Text = "预计签定合同日期"
dtgB.Col = 5: dtgB.Text = "评分" & Chr(13) & Chr(10) & "总分100"

dtgB.ColWidth(0) = 1900
dtgB.ColWidth(1) = 770
dtgB.ColWidth(2) = 1665
dtgB.ColWidth(3) = 1905
dtgB.ColWidth(4) = 2505
dtgB.ColWidth(6) = 0


End Sub

Public Sub dtgCFF()
Dim oo As Integer
dtgC.Clear
For oo = 1 To dtgC.Rows - 1
    dtgC.RowHeight(oo) = dtgP.RowHeight(0)
Next
dtgC.Row = 0
dtgC.Col = 5: dtgC.Text = "VIP"
dtgC.Col = 0: dtgC.Text = "客户名称" & Chr(13) & Chr(10) & "Customer Name"
dtgC.Col = 1: dtgC.Text = "合同类型"
dtgC.Col = 2: dtgC.Text = "满意"
dtgC.Col = 3: dtgC.Text = "有些意见"
dtgC.Col = 4: dtgC.Text = "危险"
dtgC.Col = 5: dtgC.Text = "VIP"

dtgC.ColWidth(0) = 1900
dtgC.ColWidth(1) = 770
dtgC.ColWidth(2) = 1665
dtgC.ColWidth(3) = 1905
dtgC.ColWidth(4) = 2505
dtgC.ColWidth(6) = 0


End Sub

Public Sub dtgDFF()
Dim oo As Integer
For oo = 1 To dtgD.Rows - 1
    dtgD.RowHeight(oo) = dtgD.RowHeight(0)
Next
dtgD.Clear
dtgD.Row = 0
dtgD.Col = 3: dtgD.Text = "预计金费"
dtgD.Col = 0: dtgD.Text = "项目"
dtgD.Col = 1: dtgD.Text = "地点"
dtgD.Col = 2: dtgD.Text = "计划事务"
dtgD.Col = 3: dtgD.Text = "预计金费"

dtgD.ColWidth(0) = 3180
dtgD.ColWidth(1) = 2505
dtgD.ColWidth(2) = 2895
dtgD.ColWidth(3) = 1185
dtgD.ColWidth(4) = 0
End Sub

Public Sub Bound(Uid As String, Frq As Date)
Dim tt As String
Dim Ra, Rb, RC, RD, RE, Rf, Rg, Rh, Ri, Rj, Rk
Dim La, Lb, Lc, Ld, Le, LF, Lg, Lh, Li, Lk
Dim LR As Date: Dim FR As Date
Dim oo As Integer: Dim ii As Integer
Dim DH As Single
If Month(Frq) < 4 Then
    LR = DateSerial(Year(Frq), Month(Frq), Day(Frq) + 6)
    FR = DateSerial(Year(Frq) - 1, 4, 1)
Else
    LR = DateSerial(Year(Frq), Month(Frq), Day(Frq) + 6)
    FR = DateSerial(Year(Frq), 4, 1)
End If
Me.Enabled = False
frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh
tt = "declare @gid int,@bmid tinyint,@uid nvarchar(10),@Lcuid nvarchar(10);" & _
    " select bmid,bxid,bf,bz4,bz5,bz6,gid,lc,lcuid,fwid from SalesReport where fr='" & Frq & "' and uid='" & Uid & "';" & _
    " select @gid=gid,@bmid=bmid,@uid=uid,@lcuid=lcuid from SalesReport where fr='" & Frq & "' and uid='" & Uid & "';" & _
    " select xmmc,amount,jb,jzq,khyx,pf,did from SalesReportDetail1 where gid=@gid;" & _
    " select xmmc,amount,jb,xz,yq,pf,did from SalesReportDetail2 where gid=@gid;" & _
    " select xmmc,lx,my,yj,wx,vip,did from SalesReportDetail3 where gid=@gid;" & _
    " select xmmc,adr,nr,je,did from SalesReportProject where gid=@gid;" & _
    " select bm from bm where bmid=@bmid;" & _
    " select username from worker where userid=@uid;" & _
    " select sum(amount) from SDV_ChargeA where code=@uid and billdate>='" & FR & "' and billdate<'" & LR & "';" & _
    " SELECT sum(dbo.htping1.yingfJe) " & _
        "FROM dbo.htping1 INNER JOIN dbo.htPing ON dbo.htping1.htBh = dbo.htPing.Hid " & _
        "WHERE (YEAR(dbo.htping1.rq) > 2005) AND (dbo.htPing.DelF = 1) AND (dbo.htPing.htF = 1 OR dbo.htPing.htF = 2 OR dbo.htPing.htF = 9) and dbo.htping.xuid=@uid" & _
        " and dbo.htping1.rq>='" & FR & "' and dbo.htping1.rq<'" & LR & "';" & _
         " select username from worker where userid=@lcuid;" & _
         " select trq,ywy,zn,bz,tf from pizu where bh=@gid and yid=40 order by pid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
On Error GoTo frmGZBN1E2
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
'Set mod1.HTP = mod1.HTP.NextRecordset
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RD = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RE = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rf = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rg = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rh = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Ri = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rj = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rk = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
lblBM.Caption = Rf(0, 0)
lblBM.ToolTipText = Ra(0, 0)
lblRen.Caption = Rg(0, 0)
lblRen.ToolTipText = Uid
lblWZB.Caption = Round(Rh(0, 0), 0)
lblYZ.Caption = Round(Ri(0, 0), 0) '应收帐款
LCRen = Rj(0, 0)
txtBF.Text = Ra(2, 0)
txtBxid.Text = Ra(1, 0)
txtBz4.Text = Ra(3, 0)
txtBz5.Text = Ra(4, 0)
txtBz6.Text = Ra(5, 0)
lblGid.Caption = Ra(6, 0)
Me.Lc = Ra(7, 0)
LCUid = Ra(8, 0)
Fwid = Ra(9, 0)
lblFR.Caption = Frq
lblLR.Caption = DateSerial(Year(Frq), Month(Frq), Day(Frq) + 6)
Me.Caption = "SERVICE SALES WEEKLY REPORT     " & lblFR.Caption & "  to   " & lblLR.Caption
Lb = UBound(Rb, 2) + 1
Lc = UBound(RC, 2) + 1
Ld = UBound(RD, 2) + 1
Le = UBound(RE, 2) + 1
Lk = UBound(Rk, 2) + 1
If Not (Lb > 0) Then
    Lb = 0
End If
If Not (Lc > 0) Then
    Lc = 0
End If
If Not (Ld > 0) Then
    Ld = 0
End If
If Not (Le > 0) Then
    Le = 0
End If
dtgA.Rows = Lb + 2: dtgA.ToolTipText = Lb + 2
For oo = 1 To Lb + 1
    dtgA.Row = oo
    For ii = 0 To 6
        dtgA.Col = ii
        dtgA.Text = Rb(ii, oo - 1)
        If ii <> 1 Then
            DH = 255 * mod1.HH(dtgA.Text, UpInt(dtgA.CellWidth / 200))
            If DH > dtgA.RowHeight(dtgA.Row) Then
                dtgA.RowHeight(dtgA.Row) = DH
            End If
        End If
    Next
Next
dtgB.Rows = Lc + 2: dtgB.ToolTipText = Lc + 2
For oo = 1 To Lc + 1
    dtgB.Row = oo
    For ii = 0 To 6
        dtgB.Col = ii
        dtgB.Text = RC(ii, oo - 1)
        If ii <> 1 Then
            DH = 255 * mod1.HH(dtgB.Text, UpInt(dtgB.CellWidth / 200))
            If DH > dtgB.RowHeight(dtgB.Row) Then
                dtgB.RowHeight(dtgB.Row) = DH
            End If
        End If
    Next
Next
dtgC.Rows = Ld + 2: dtgC.ToolTipText = Ld + 2
For oo = 1 To Ld + 1
    dtgC.Row = oo
    For ii = 0 To 6
        dtgC.Col = ii
        dtgC.Text = RD(ii, oo - 1)
        DH = 255 * mod1.HH(dtgC.Text, UpInt(dtgC.CellWidth / 200))
        If DH > dtgC.RowHeight(dtgC.Row) Then
            dtgC.RowHeight(dtgC.Row) = DH
        End If
    Next
Next
dtgD.Rows = Le + 2: dtgD.ToolTipText = Le + 2
For oo = 1 To Le + 1
    dtgD.Row = oo
    For ii = 0 To 4
        dtgD.Col = ii
        dtgD.Text = RE(ii, oo - 1)
        DH = 255 * mod1.HH(dtgD.Text, UpInt(dtgD.CellWidth / 200))
        If DH > dtgD.RowHeight(dtgD.Row) Then
            dtgD.RowHeight(dtgD.Row) = DH
        End If
    Next
Next
If Me.Lc = 1 Or Me.Lc = 100 Or Me.Lc = 0 Then
    lblTx.Visible = False
Else
    lblTx.Caption = "流程至:" & LCRen
    lblTx.Visible = True
End If
Call dtgPFF
For oo = 1 To Lk
    dtgP.Row = oo
    For ii = 0 To 5
        dtgP.Col = ii
        dtgP.Text = Rk(ii, oo - 1)

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
For oo = 1 To Lk
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
frmWait.Visible = False
Me.Enabled = True
Me.ZOrder 0
Exit Sub
frmGZBN1E2:
MsgBox "出错!程序将关闭!"
End
End Sub

Public Sub GDui() '格式刷新
Dim oo As Integer
Dim ii As Integer
Dim rr As Integer
Dim Tm0 As String
Dim Tm5 As Single
dtgN.Clear
dtgN.Rows = dtgA.Rows
dtgN.Cols = dtgA.Cols
rr = 0
For oo = 0 To dtgA.Rows - 2
    dtgA.Row = oo
    dtgN.Row = rr
    dtgA.Col = 0: Tm0 = dtgA.Text
    dtgA.Col = 5: Tm5 = Val(dtgA.Text)
    If Not (Tm0 = "" Or Tm5 = 0) Or oo = 0 Then
        For ii = 0 To 6
            dtgA.Col = ii: dtgN.Col = ii
            dtgN.Text = dtgA.Text

        Next
        rr = rr + 1
    Else
        dtgA.Col = 0: dtgA.Text = ""
        dtgA.Col = 5: dtgA.Text = ""
    End If
Next
dtgN.Rows = rr + 1
dtgA.Rows = rr + 1
dtgA.Clear
For oo = 0 To dtgN.Rows - 2
    dtgN.Row = oo: dtgA.Row = oo
    For ii = 0 To 6
        dtgN.Col = ii: dtgA.Col = ii
        dtgA.Text = dtgN.Text
    Next
Next
dtgA.ToolTipText = rr + 2

dtgN.Clear
dtgN.Rows = dtgB.Rows
dtgN.Cols = dtgB.Cols
rr = 0
For oo = 0 To dtgB.Rows - 2
    dtgB.Row = oo
    dtgN.Row = rr
    dtgB.Col = 0: Tm0 = dtgB.Text
    dtgB.Col = 5: Tm5 = Val(dtgB.Text)
    If Not (Tm0 = "" Or Tm5 = 0) Or oo = 0 Then
        For ii = 0 To 6
            dtgB.Col = ii: dtgN.Col = ii
            dtgN.Text = dtgB.Text

        Next
        rr = rr + 1
    Else
        dtgB.Col = 0: dtgB.Text = ""
        dtgB.Col = 5: dtgB.Text = ""
    End If
Next
dtgN.Rows = rr + 1
dtgB.Rows = rr + 1
dtgB.Clear
For oo = 0 To dtgN.Rows - 2
    dtgN.Row = oo: dtgB.Row = oo
    For ii = 0 To 6
        dtgN.Col = ii: dtgB.Col = ii
        dtgB.Text = dtgN.Text
    Next
Next
dtgB.ToolTipText = rr + 2

dtgN.Clear
dtgN.Rows = dtgC.Rows
dtgN.Cols = dtgC.Cols
rr = 0
For oo = 0 To dtgC.Rows - 2
    dtgC.Row = oo
    dtgN.Row = rr
    dtgC.Col = 0: Tm0 = dtgC.Text
    dtgC.Col = 5: Tm5 = Val(dtgC.Text)
    If Not (Tm0 = "" Or Tm5 = 0) Or oo = 0 Then
        For ii = 0 To 6
            dtgC.Col = ii: dtgN.Col = ii
            dtgN.Text = dtgC.Text

        Next
        rr = rr + 1
    Else
        dtgC.Col = 0: dtgC.Text = ""
        dtgC.Col = 5: dtgC.Text = ""
    End If
Next
dtgN.Rows = rr + 1
dtgC.Rows = rr + 1
dtgC.Clear
For oo = 0 To dtgN.Rows - 2
    dtgN.Row = oo: dtgC.Row = oo
    For ii = 0 To 6
        dtgN.Col = ii: dtgC.Col = ii
        dtgC.Text = dtgN.Text
    Next
Next
dtgC.ToolTipText = rr + 2

dtgN.Clear
dtgN.Rows = dtgD.Rows
dtgN.Cols = dtgD.Cols
rr = 0
For oo = 0 To dtgD.Rows - 2
    dtgD.Row = oo
    dtgN.Row = rr
    dtgD.Col = 0: Tm0 = dtgD.Text
    dtgD.Col = 3: Tm5 = Val(dtgD.Text)
    If Not (Tm0 = "" Or Tm5 = 0) Or oo = 0 Then
        For ii = 0 To 4
            dtgD.Col = ii: dtgN.Col = ii
            dtgN.Text = dtgD.Text

        Next
        rr = rr + 1
    Else
        dtgD.Col = 0: dtgD.Text = ""
        dtgD.Col = 3: dtgD.Text = ""
    End If
Next
dtgN.Rows = rr + 1
dtgD.Rows = rr + 1
dtgD.Clear
For oo = 0 To dtgN.Rows - 2
    dtgN.Row = oo: dtgD.Row = oo
    For ii = 0 To 4
        dtgN.Col = ii: dtgD.Col = ii
        dtgD.Text = dtgN.Text
    Next
Next
dtgD.ToolTipText = rr + 2
End Sub
