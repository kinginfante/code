VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGZDJY 
   Caption         =   "工作单检验"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton Command2 
      Caption         =   "查询"
      Height          =   315
      Left            =   14280
      TabIndex        =   27
      Top             =   7440
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   285
      Left            =   14280
      TabIndex        =   26
      Top             =   6780
      Width           =   795
   End
   Begin VB.CommandButton cmdRen 
      Caption         =   "业务员"
      Height          =   315
      Left            =   10740
      TabIndex        =   7
      Top             =   7440
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   11850
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6780
      Width           =   2265
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   11850
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   6300
      Width           =   2265
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "打    开"
      Height          =   345
      Left            =   10350
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   4845
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "导航"
      Height          =   585
      Left            =   14520
      Picture         =   "frmGZDJY.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8550
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgGzd 
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   16113
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame frmAdd 
      Height          =   5595
      Left            =   10290
      TabIndex        =   9
      Top             =   300
      Width           =   5025
      Begin MSDataListLib.DataList liXmmc 
         Height          =   2790
         Left            =   1140
         TabIndex        =   31
         Top             =   1470
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   4921
         _Version        =   393216
      End
      Begin VB.TextBox txtXmmc 
         Height          =   270
         Left            =   3030
         TabIndex        =   30
         Top             =   4320
         Width           =   1545
      End
      Begin VB.CommandButton cmdZJ 
         Caption         =   "质量监督部"
         Height          =   315
         Left            =   1830
         TabIndex        =   28
         Top             =   3630
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加"
         Height          =   285
         Left            =   270
         TabIndex        =   22
         Top             =   4500
         Width           =   1035
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         Height          =   285
         Left            =   270
         TabIndex        =   21
         Top             =   4800
         Width           =   1005
      End
      Begin VB.OptionButton optLx 
         Caption         =   "冷水机组巡视检修工作报告（单）"
         Height          =   315
         Index           =   1
         Left            =   1740
         TabIndex        =   20
         Top             =   420
         Width           =   3015
      End
      Begin VB.OptionButton optLx 
         Caption         =   "热泵机组巡视检修工作报告（单）"
         Height          =   315
         Index           =   2
         Left            =   1740
         TabIndex        =   19
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton optLx 
         Caption         =   "冷水机组年度检修工作报告（单）"
         Height          =   315
         Index           =   3
         Left            =   1740
         TabIndex        =   18
         Top             =   1245
         Width           =   3015
      End
      Begin VB.OptionButton optLx 
         Caption         =   "热泵机组年度检修工作报告（单）"
         Height          =   315
         Index           =   4
         Left            =   1740
         TabIndex        =   17
         Top             =   1665
         Width           =   3015
      End
      Begin VB.OptionButton optLx 
         Caption         =   "应急维修工作报告（单）"
         Height          =   315
         Index           =   5
         Left            =   1740
         TabIndex        =   16
         Top             =   2085
         Width           =   3015
      End
      Begin VB.OptionButton optLx 
         Caption         =   "机组大修工作报告（单）"
         Height          =   315
         Index           =   6
         Left            =   1740
         TabIndex        =   15
         Top             =   2490
         Width           =   3015
      End
      Begin VB.OptionButton optLx 
         Caption         =   "施工工作报告（单）"
         Height          =   315
         Index           =   7
         Left            =   1740
         TabIndex        =   14
         Top             =   2910
         Width           =   3015
      End
      Begin VB.OptionButton optLx 
         Caption         =   "工程部维修质量监督报告（单）"
         Height          =   315
         Index           =   8
         Left            =   1740
         TabIndex        =   13
         Top             =   3330
         Width           =   3015
      End
      Begin VB.CommandButton cmdFw 
         Caption         =   "选择人员"
         Height          =   315
         Left            =   1830
         TabIndex        =   12
         Top             =   3930
         Width           =   1095
      End
      Begin VB.TextBox txtBh 
         Height          =   270
         Left            =   3030
         TabIndex        =   11
         Top             =   4740
         Width           =   1545
      End
      Begin VB.CheckBox chkHGF 
         Caption         =   "不合格工作单"
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   2310
         TabIndex        =   10
         Top             =   5100
         Width           =   2235
      End
      Begin VB.Label Label5 
         Caption         =   "按项目名称选择业务员"
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   1950
         TabIndex        =   29
         Top             =   4290
         Width           =   975
      End
      Begin VB.Shape Shape1 
         Height          =   5325
         Left            =   90
         Shape           =   4  'Rounded Rectangle
         Top             =   150
         Width           =   4845
      End
      Begin VB.Label Label1 
         Caption         =   "工作单类型"
         Height          =   285
         Left            =   300
         TabIndex        =   25
         Top             =   390
         Width           =   1185
      End
      Begin VB.Label lblFw 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3060
         TabIndex        =   24
         Top             =   3960
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "编号"
         Height          =   255
         Left            =   2280
         TabIndex        =   23
         Top             =   4770
         Width           =   495
      End
   End
   Begin VB.Label lblRen 
      Caption         =   "Label5"
      Height          =   255
      Left            =   11910
      TabIndex        =   8
      Top             =   7470
      Width           =   2205
   End
   Begin VB.Line Line1 
      X1              =   10410
      X2              =   15225
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label Label4 
      Caption         =   "值"
      Height          =   225
      Left            =   10710
      TabIndex        =   5
      Top             =   6810
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "查询方式:"
      Height          =   225
      Left            =   10680
      TabIndex        =   3
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      Height          =   2055
      Left            =   10410
      Shape           =   4  'Rounded Rectangle
      Top             =   6060
      Width           =   4815
   End
End
Attribute VB_Name = "frmGZDJY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoGZD As ADODB.Recordset
Dim adoXmmc As ADODB.Recordset
Dim Fl As Integer
Dim Gid As Long

Private Sub cmdAdd_Click()
On Error Resume Next
Dim tt As String
Dim hgF As Integer
'检查有无重复单子
tt = "select bh from newGzd where bh='" & txtBh.Text & "'"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.HTP.RecordCount > 0 Then
    MsgBox "不能输入重复单子！"
    Exit Sub
End If

'基础发布
    tt = "select userid from worker where username='" & lblFw.Caption & "'"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    lblFw.ToolTipText = mod1.HTP.Fields("userid").Value

    '重置类型框
    Select Case Left(txtBh.Text, 1)
    Case 1
        optLx(5).Value = True
    Case 2
        optLx(6).Value = True
    Case 3
        optLx(1).Value = True
    Case 4
        optLx(2).Value = True
    Case 5
        optLx(4).Value = True
    Case 6
        optLx(3).Value = True
    Case 7
        optLx(7).Value = True
    Case 8
        optLx(8).Value = True
    End Select

    If txtBh.Text = "" Or lblFw.Caption = "" Or Fl = 0 Then
        Exit Sub
    End If
    If chkHGF.Value = 1 Then
        hgF = 0
    Else
        hgF = 1
    End If
    tt = "insert newGzd (ywy,uid,bh,lrq,fl,hgf) values ('" & lblFw.Caption & "','" & lblFw.ToolTipText & "','" & txtBh.Text & "','" & DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda)) & "'," & Fl & "," & hgF & ")"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockBatchOptimistic, adCmdText
    adoGZD.Requery
    Set dtgGzd.DataSource = adoGZD
    chkHGF.Value = 0

End Sub

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
End Sub

Private Sub cmdDel_Click()
Dim ii As Integer
Dim Xmmc As String
Dim tt As String
On Error Resume Next
Xmmc = ""
dtgGzd.Col = 9
Xmmc = dtgGzd.Text
If Xmmc <> "" Then
    Exit Sub
End If
dtgGzd.Col = 5
Gid = dtgGzd.Text
ii = MsgBox("是否删除此条记录?", vbInformation + vbYesNo, "询问")
If ii = vbYes Then
    tt = "delete from newGzd where gid=" & Gid
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockBatchOptimistic, adCmdText
    adoGZD.Requery
    Set dtgGzd.DataSource = adoGZD
End If
End Sub

Private Sub cmdFw_Click()
Set Ren.XForm = New frmGZDJY
Call mod1.RenXz("frmGZDJY", Me, 0)
End Sub

Private Sub cmdZJ_Click()
lblFw.Caption = "李铭"
lblFw.ToolTipText = "HM361"
End Sub

Private Sub dtgGzd_Click()
liXmmc.Visible = False
End Sub

Private Sub dtgGzd_DblClick()
Static Px As Boolean

If dtgGzd.Row = 1 Then
    If Px = True Then
        dtgGzd.Sort = 2
        Px = False
    Else
        dtgGzd.Sort = 1
        Px = True
    End If
'Else
'    MsgBox MGa.ColData(1)
End If
End Sub

Private Sub dtgGzd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Zf As Boolean
If Button <> 2 Then Exit Sub
If Zf = False Then
        htBrowG.dtgBr.FixedRows = 0
        htBrowG.dtgBr.MergeCol(1) = True
        htBrowG.dtgBr.MergeCol(2) = True
        htBrowG.dtgBr.MergeCol(3) = True
        htBrowG.dtgBr.MergeCol(4) = True
        htBrowG.dtgBr.MergeCol(7) = True
        htBrowG.dtgBr.MergeCol(13) = True
        htBrowG.dtgBr.MergeCells = 0
        htBrowG.dtgBr.FixedRows = 1
        Zf = True
Else
        htBrowG.dtgBr.FixedRows = 0
        htBrowG.dtgBr.MergeCol(1) = True
        htBrowG.dtgBr.MergeCol(2) = True
        htBrowG.dtgBr.MergeCol(3) = True
        htBrowG.dtgBr.MergeCol(4) = True
        htBrowG.dtgBr.MergeCol(7) = True
        htBrowG.dtgBr.MergeCol(13) = True
        htBrowG.dtgBr.MergeCells = 3
        htBrowG.dtgBr.FixedRows = 1
        Zf = False
End If
End Sub

Private Sub Form_Click()
liXmmc.Visible = False
End Sub

Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
Set adoGZD = New ADODB.Recordset
dtgGzd.ColWidth(0) = 300
dtgGzd.ColWidth(1) = 1000
dtgGzd.ColWidth(4) = 3000
dtgGzd.ColWidth(5) = 0
dtgGzd.ColWidth(6) = 0
dtgGzd.ColWidth(7) = 0
dtgGzd.ColWidth(8) = 0
dtgGzd.ColWidth(9) = 2500
dtgGzd.ColWidth(10) = 0
dtgGzd.ColWidth(11) = 0
frmAdd.BorderStyle = 0
Set adoXmmc = New ADODB.Recordset
liXmmc.Visible = False
End Sub

Private Sub frmAdd_Click()
liXmmc.Visible = False
End Sub

Private Sub liXmmc_DblClick()
On Error Resume Next
If adoXmmc.RecordCount > 0 Then
    lblFw.Caption = liXmmc.BoundText
End If
    liXmmc.Visible = False
    txtXMMC.Text = ""
End Sub


Private Sub optLx_Click(Index As Integer)
Fl = Index
End Sub

Private Sub txtBh_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim hgF As Integer
Dim tt As String

If KeyCode = 13 Then
'基础发布
    tt = "select userid from worker where username='" & lblFw.Caption & "'"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    lblFw.ToolTipText = mod1.HTP.Fields("userid").Value

    '重置类型框
    Select Case Left(txtBh.Text, 1)
    Case 1
        optLx(5).Value = True
    Case 2
        optLx(6).Value = True
    Case 3
        optLx(1).Value = True
    Case 4
        optLx(2).Value = True
    Case 5
        optLx(4).Value = True
    Case 6
        optLx(3).Value = True
    Case 7
        optLx(7).Value = True
    Case 8
        optLx(8).Value = True
    End Select

    If txtBh.Text = "" Or lblFw.Caption = "" Or Fl = 0 Then
        Exit Sub
    End If
    If chkHGF.Value = 1 Then
        hgF = 0
    Else
        hgF = 1
    End If
    tt = "insert newGzd (ywy,uid,bh,lrq,fl,hgf) values ('" & lblFw.Caption & "','" & lblFw.ToolTipText & "','" & txtBh.Text & "','" & mod1.DQda & "'," & Fl & "," & hgF & ")"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockBatchOptimistic, adCmdText
    adoGZD.Requery
    Set dtgGzd.DataSource = adoGZD
    chkHGF.Value = 0
End If
End Sub

Private Sub txtXmmc_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tt As String
On Error Resume Next
If KeyCode = 13 And Trim(txtXMMC.Text) <> "" Then
    tt = "select xmmc,ywy from xmzl where xmmc like '%" & txtXMMC.Text & "%'"
    adoXmmc.Close
    '基础发布
    adoXmmc.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set liXmmc.RowSource = adoXmmc
    liXmmc.ListField = "xmmc"
    liXmmc.BoundColumn = "ywy"
    liXmmc.Visible = True
End If
End Sub


