VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGZDBR 
   Caption         =   "工作单查询"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Frame frmFw 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   765
      Left            =   9480
      TabIndex        =   12
      Top             =   8370
      Width           =   2355
      Begin VB.CommandButton cmdZJ 
         Caption         =   "质量监督部"
         Height          =   285
         Left            =   1050
         TabIndex        =   22
         Top             =   150
         Width           =   1245
      End
      Begin VB.CommandButton cmdFw 
         Caption         =   "选择业务员"
         Height          =   285
         Left            =   0
         TabIndex        =   13
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label lblFw 
         Height          =   225
         Left            =   30
         TabIndex        =   14
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   705
      Left            =   10620
      TabIndex        =   15
      Top             =   8490
      Visible         =   0   'False
      Width           =   3945
      Begin VB.CommandButton cmdV 
         Caption         =   "查询"
         Height          =   285
         Left            =   3510
         TabIndex        =   19
         Top             =   30
         Width           =   825
      End
      Begin VB.TextBox txtW 
         Height          =   285
         Left            =   2250
         TabIndex        =   18
         Top             =   0
         Width           =   1155
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "frmGZDBR.frx":0000
         Left            =   810
         List            =   "frmGZDBR.frx":000A
         TabIndex        =   17
         Text            =   "编号"
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton cmdAll2 
         Caption         =   "全  部"
         Height          =   285
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "值"
         Height          =   315
         Left            =   1830
         TabIndex        =   21
         Top             =   30
         Width           =   315
      End
      Begin VB.Label Label6 
         Caption         =   "查询方式"
         Height          =   285
         Left            =   60
         TabIndex        =   20
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "全  部"
      Height          =   285
      Left            =   90
      TabIndex        =   11
      Top             =   8880
      Width           =   9405
   End
   Begin VB.CommandButton cmdOpen2 
      Caption         =   "打    开"
      Height          =   255
      Left            =   12240
      TabIndex        =   10
      Top             =   30
      Width           =   2955
   End
   Begin VB.CommandButton cmdOpen1 
      Caption         =   "打    开"
      Height          =   285
      Left            =   3720
      TabIndex        =   9
      Top             =   30
      Width           =   5775
   End
   Begin VB.CommandButton cmdREF 
      Caption         =   "查询"
      Height          =   285
      Left            =   8250
      TabIndex        =   8
      Top             =   8550
      Width           =   1245
   End
   Begin VB.TextBox txtZ 
      Height          =   285
      Left            =   4530
      TabIndex        =   6
      Top             =   8520
      Width           =   3555
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "frmGZDBR.frx":001E
      Left            =   1200
      List            =   "frmGZDBR.frx":002B
      TabIndex        =   5
      Text            =   "编号"
      Top             =   8520
      Width           =   2715
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "返回"
      Height          =   555
      Left            =   14580
      Picture         =   "frmGZDBR.frx":0045
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8580
      Width           =   645
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgW 
      Height          =   7995
      Left            =   9510
      TabIndex        =   2
      Top             =   300
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   14102
      _Version        =   393216
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgY 
      Height          =   7995
      Left            =   -30
      TabIndex        =   23
      Top             =   300
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   14102
      _Version        =   393216
      FillStyle       =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label4 
      Caption         =   "值"
      Height          =   315
      Left            =   4110
      TabIndex        =   7
      Top             =   8550
      Width           =   315
   End
   Begin VB.Label Label3 
      Caption         =   "查询方式"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   8520
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "未完成工作单"
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   10260
      TabIndex        =   1
      Top             =   60
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "工作单记录"
      ForeColor       =   &H8000000D&
      Height          =   225
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   1365
   End
End
Attribute VB_Name = "frmGZDBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoY As ADODB.Recordset '已完成的工作单
Public adoW As ADODB.Recordset '未完成的工作单

Private Sub cmdAll_Click()
On Error Resume Next
Dim tt As String
tt = "select 检验日期,编号,工作单类型,gid,fl FROM gzdView where trq is null and 业务员='" & lblFw.Caption & "' and uid='" & lblFw.ToolTipText & "' order by 检验日期"
frmGZDBR.adoW.Close
frmGZDBR.adoW.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgW.DataSource = frmGZDBR.adoW

tt = "Select 检验日期,编号,工作单类型,业务员,gid,uid,qy,trq,项目名称,日期,fl,合格否 from gzdView where not(trq is null) and 业务员='" & lblFw.Caption & "' and uid='" & lblFw.ToolTipText & "' order by gid desc"
frmGZDBR.adoY.Close
frmGZDBR.adoY.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgY.DataSource = frmGZDBR.adoY
frmGZDBR.dtgY.Row = 1

'Set frmGZDBR.dtgY.DataSource = frmGZDBR.adoW
End Sub

Private Sub cmdAll2_Click()
On Error Resume Next
Dim tt As String
tt = "select 检验日期,编号,工作单类型,gid,fl FROM gzdView where trq is null and 业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "' order by 检验日期"
frmGZDBR.adoW.Close
frmGZDBR.adoW.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgW.DataSource = frmGZDBR.adoW
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
frmZu.ZOrder 0
End Sub

Private Sub cmdFw_Click()
Set Ren.XForm = New frmGZDBR
Call mod1.RenXz("frmGZDBR", Me, 0)
End Sub

Private Sub cmdOpen1_Click()
On Error Resume Next
Dim Gid As Long
Dim Fl As Integer
dtgY.Col = 5
Gid = dtgY.Text
dtgY.Col = 11
Fl = dtgY.Text
'Select Case Fl
'Case 1
'    Fl = 5
'Case 2
'    Fl = 6
'Case 3
'    Fl = 1
'Case 4
'    Fl = 2
'Case 5
'    Fl = 4
'Case 6
'    Fl = 3
'Case 7
'    Fl = 7
'Case 8
'    Fl = 8
'End Select
Select Case Fl
Case 1
    NewGZD1.Show
    NewGZD1.ZOrder 0
    Call modGZD.gzd1Qing
    Call modGZD.gzd1Bound(Gid)
Case 2
    NewGzd2.Show
    NewGzd2.ZOrder 0
    Call modGZD.gzd2Qing
    Call modGZD.gzd2Bound(Gid)
Case 3
    NewGzd3.Show
    NewGzd3.ZOrder 0
    Call modGZD.gzd3Qing
    Call modGZD.gzd3Bound(Gid)
Case 4
    NewGzd4.Show
    NewGzd4.ZOrder 0
    Call modGZD.gzd4Qing
    Call modGZD.gzd4Bound(Gid)
Case 5
    NewGzd5.Show
    NewGzd5.ZOrder 0
    Call modGZD.gzd5Qing
    Call modGZD.gzd5Bound(Gid)
Case 6
    NewGzd6.Show
    NewGzd6.ZOrder 0
    Call modGZD.gzd6Qing
    Call modGZD.gzd6Bound(Gid)
Case 7
    NewGzd7.Show
    NewGzd7.ZOrder 0
    Call modGZD.gzd7Qing
    Call modGZD.gzd7Bound(Gid)
Case 8
    NewGZD8.Show
    NewGZD8.ZOrder 0
    'Call modGZD.gzd8Qing
    'Call modGZD.gzd8Bound(Gid)
End Select
    frmGZDBR.Enabled = False
End Sub

Private Sub cmdOpen2_Click()
On Error Resume Next
Dim Gid As Long
Dim Fl As Integer
dtgW.Col = 4
Gid = dtgW.Text
dtgW.Col = 5
Fl = dtgW.Text
Select Case Fl
Case 1
    NewGZD1.Show
    NewGZD1.ZOrder 0
    Call modGZD.gzd1Qing
    Call modGZD.gzd1Bound(Gid)
Case 2
    NewGzd2.Show
    NewGzd2.ZOrder 0
    Call modGZD.gzd2Qing
    Call modGZD.gzd2Bound(Gid)
Case 3
    NewGzd3.Show
    NewGzd3.ZOrder 0
    Call modGZD.gzd3Qing
    Call modGZD.gzd3Bound(Gid)
Case 4
    NewGzd4.Show
    NewGzd4.ZOrder 0
    Call modGZD.gzd4Qing
    Call modGZD.gzd4Bound(Gid)
Case 5
    NewGzd5.Show
    NewGzd5.ZOrder 0
    Call modGZD.gzd5Qing
    Call modGZD.gzd5Bound(Gid)
Case 6
    NewGzd6.Show
    NewGzd6.ZOrder 0
    Call modGZD.gzd6Qing
    Call modGZD.gzd6Bound(Gid)
Case 7
    NewGzd7.Show
    NewGzd7.ZOrder 0
    Call modGZD.gzd7Qing
    Call modGZD.gzd7Bound(Gid)
Case 8
    NewGZD8.Show
    NewGZD8.ZOrder 0
    Call modGZD.gzd8Qing
    Call modGZD.gzd8Bound(Gid)
End Select
    frmGZDBR.Enabled = False

End Sub

Private Sub cmdRef_Click()
On Error Resume Next
Dim tt As String
If comLx.Text = "编号" Then
    If mod1.DName = "李铭" Or mod1.DName = "张寅" Or mod1.DName = "倪旭" Or mod1.DName = "宋晓炯" Or mod1.DName = "钱亘" Then
        tt = "Select * from gzdView where not(trq is null) and 编号=" & Val(txtZ.Text) & "  order by gid desc"
    ElseIf mod1.KhK = 1 Then
        tt = "Select * from gzdView where not(trq is null) and 编号=" & Val(txtZ.Text) & " and bm='" & mod1.BM & "' order by gid desc"
    Else
        tt = "Select * from gzdView where not(trq is null) and 编号=" & Val(txtZ.Text) & " and 业务员='" & mod1.DName & "' order by gid desc"
    End If

ElseIf comLx.Text = "项目名称" Then
'    If mod1.DName = "钱亘" Or mod1.DName = "张寅" Or mod1.DName = "王卫卫" Or mod1.DName = "倪旭" Or mod1.DName = "宋晓炯" Then
'        tt = "Select * from gzdView where not(trq is null) and 项目名称 like '%" & txtZ.Text & "%' and 业务员='" & mod1.DName & "' order by gid desc"
'    ElseIf mod1.KhK = 1 Then
'        tt = "Select * from gzdView where not(trq is null) and 项目名称 like '%" & txtZ.Text & "%' and bm='" & mod1.DName & "' order by gid desc"
'    ElseIf mod1.DName = "钱亘" Or mod1.DName = "张寅" Or mod1.DName = "王卫卫" Or mod1.DName = "倪旭" Or mod1.DName = "宋晓炯" Then
'        tt = "Select * from gzdView where not(trq is null) and 项目名称 like '%" & txtZ.Text & "%'  order by gid desc"
'    End If

    If mod1.DName = "李铭" Or mod1.DName = "张寅" Or mod1.DName = "倪旭" Or mod1.DName = "宋晓炯" Or mod1.DName = "钱亘" Then
        tt = "Select * from gzdView where not(trq is null) and 项目名称 like '%" & txtZ.Text & "%'  order by gid desc"
    ElseIf mod1.KhK = 1 Then
        tt = "Select * from gzdView where not(trq is null) and 项目名称 like '%" & txtZ.Text & "%' and bm='" & mod1.BM & "' order by gid desc"
    Else
        tt = "Select * from gzdView where not(trq is null) and 项目名称 like '%" & txtZ.Text & "%' and 业务员='" & mod1.DName & "' order by gid desc"
    End If
ElseIf comLx.Text = "类别" Then
    If mod1.DName = "李铭" Or mod1.DName = "张寅" Or mod1.DName = "倪旭" Or mod1.DName = "宋晓炯" Or mod1.DName = "钱亘" Then
        tt = "Select * from gzdView where not(trq is null) and 工作单类型 like '%" & txtZ.Text & "%'  order by gid desc"
    ElseIf mod1.KhK = 1 Then
        tt = "Select * from gzdView where not(trq is null) and 工作单类型 like '%" & txtZ.Text & "%' and bm='" & mod1.BM & "' order by gid desc"
    Else
        tt = "Select * from gzdView where not(trq is null) and 工作单类型 like '%" & txtZ.Text & "%' and 业务员='" & mod1.DName & "' order by gid desc"
    End If
End If
frmGZDBR.adoY.Close
frmGZDBR.adoY.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgY.DataSource = frmGZDBR.adoY

If comLx.Text = "编号" Then
    tt = "Select 检验日期,编号,工作单类型,gid,fl from gzdView where trq is null and 编号=" & Val(txtZ.Text) & " order by gid desc"
ElseIf comLx.Text = "项目名称" Then
    tt = ""
   ' tt = "Select * from gzdView where trq is null and 项目名称 like '%" & txtZ.Text & "%' and 业务员='" & mod1.DName & "' order by gid desc"
End If
frmGZDBR.adoW.Close
frmGZDBR.adoW.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgW.DataSource = frmGZDBR.adoW

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdV_Click()
On Error Resume Next
Dim tt As String

tt = "Select  检验日期,编号,工作单类型,gid,fl  from gzdView where trq is null and 编号=" & Val(txtW.Text) & " order by gid desc"

frmGZDBR.adoW.Close
frmGZDBR.adoW.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgW.DataSource = frmGZDBR.adoW
End Sub


Private Sub cmdZJ_Click()
Dim tt As String
On Error Resume Next
lblFw.Caption = "钱亘"
lblFw.ToolTipText = "HM152"

tt = "Select * from gzdView where not(trq is null) and 业务员='" & lblFw.Caption & "' and uid='" & lblFw.ToolTipText & "' order by gid desc"
frmGZDBR.adoY.Close
frmGZDBR.adoY.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgY.DataSource = frmGZDBR.adoY

tt = "select 检验日期,编号,工作单类型,gid,fl FROM gzdView where trq is null and 业务员='" & lblFw.Caption & "' and uid='" & lblFw.ToolTipText & "' order by 检验日期"
frmGZDBR.adoW.Close
frmGZDBR.adoW.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgW.DataSource = frmGZDBR.adoW
End Sub

Private Sub dtgW_DblClick()
Static Px As Boolean

If dtgW.Row = 1 Then
    If Px = True Then
        dtgW.Sort = 2
        Px = False
    Else
        dtgW.Sort = 1
        Px = True
    End If
'Else
'    MsgBox MGa.ColData(1)
End If
End Sub

Private Sub dtgY_DblClick()
Static Px As Boolean

If dtgY.Row = 1 Then
    If Px = True Then
        dtgY.Sort = 2
        Px = False
    Else
        dtgY.Sort = 1
        Px = True
    End If
'Else
'    MsgBox MGa.ColData(1)
End If
End Sub


Private Sub Form_Load()
Set adoY = New ADODB.Recordset
Set adoW = New ADODB.Recordset
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
dtgW.ColWidth(0) = 0
dtgW.ColWidth(4) = 0
dtgW.ColWidth(5) = 0
dtgW.ColWidth(3) = 3000

dtgY.ColWidth(0) = 0
dtgY.ColWidth(1) = 1000
dtgY.ColWidth(2) = 800
dtgY.ColWidth(3) = 3000
dtgY.ColWidth(9) = 2700
dtgY.ColWidth(4) = 0
dtgY.ColWidth(5) = 0
dtgY.ColWidth(6) = 0
dtgY.ColWidth(7) = 0
dtgY.ColWidth(8) = 0
dtgY.ColWidth(11) = 0
dtgY.ColWidth(12) = 600
dtgY.ColWidth(13) = 0
End Sub
