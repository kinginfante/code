VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGxBiao 
   Caption         =   "询价记录表"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdNew 
      Caption         =   "新建询价单"
      Height          =   735
      Left            =   12060
      TabIndex        =   10
      Top             =   5820
      Width           =   1725
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   555
      Left            =   14580
      Picture         =   "frmGxBiao1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8610
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Caption         =   "零配件数据库"
      Height          =   4455
      Left            =   7110
      TabIndex        =   1
      Top             =   30
      Width           =   8145
      Begin VB.CommandButton cmdDunham 
         Height          =   675
         Left            =   4890
         Picture         =   "frmGxBiao1.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1890
         Width           =   1815
      End
      Begin VB.CommandButton cmdKl 
         BackColor       =   &H80000009&
         Height          =   645
         Left            =   4890
         Picture         =   "frmGxBiao1.frx":0DDB
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   630
         Width           =   1815
      End
      Begin VB.CommandButton cmdYk 
         Caption         =   "约克"
         Height          =   585
         Left            =   4890
         TabIndex        =   6
         Top             =   1290
         Width           =   1815
      End
      Begin VB.CommandButton cmdTl 
         Caption         =   "特灵"
         Height          =   585
         Left            =   4890
         TabIndex        =   5
         Top             =   2580
         Width           =   1785
      End
      Begin VB.CommandButton cmdMk 
         Caption         =   "麦克威尔"
         Height          =   615
         Left            =   4890
         TabIndex        =   4
         Top             =   3180
         Width           =   1785
      End
      Begin VB.CommandButton cmdPj 
         Caption         =   "常用配件"
         Height          =   585
         Left            =   1890
         TabIndex        =   3
         Top             =   3210
         Width           =   1665
      End
      Begin VB.CommandButton cmdZlG 
         Caption         =   "制冷剂"
         Height          =   555
         Left            =   1890
         TabIndex        =   2
         Top             =   2610
         Width           =   1665
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   16113
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmGxBiao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
frmGxBiao.Visible = False
frmZu.Enabled = True
End Sub

Private Sub cmdDunham_Click()
Set frmLingjian.LpXh = New ADODB.Recordset
Dim tt As String
Dim oo As Integer
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
End If
'MsgBox "您好!目前顿汉布什的进价略有差异，我正在修改之中，具体的成本价格今年仍按以前的计算。其他品牌没有变化。谢谢  小张 分机111"
frmLingjian.Caption = "顿汉布什"
frmLingjian.Show

For oo = frmLingjian.comJzXh.ListCount - 1 To 0 Step -1
    frmLingjian.comJzXh.RemoveItem oo
Next

tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
frmLingjian.LpXh.Close
frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
frmLingjian.dtgView.Columns(5).Caption = "库存价"
If mod1.VLP = 1 Then
    frmLingjian.dtgView.Columns("库存价").Visible = False
ElseIf mod1.VLP = 2 Then
    frmLingjian.dtgView.Columns("库存价").Visible = True
ElseIf mod1.VLP = 3 Then
    frmLingjian.dtgView.Columns("库存价").Visible = True
End If
    Set frmLingjian.dtgView.DataSource = Nothing
If mod1.VLP = 3 Then
    frmLingjian.cmdKq.Visible = True
Else
    frmLingjian.cmdKq.Visible = False
End If
cmdGx.Enabled = False
End Sub

Private Sub cmdKl_Click()
Set frmLingjian.LpXh = New ADODB.Recordset
Dim tt As String
Dim oo As Integer
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
End If
frmLingjian.Caption = "开利"
frmLingjian.Show

For oo = frmLingjian.comJzXh.ListCount - 1 To 0 Step -1
    frmLingjian.comJzXh.RemoveItem oo
Next

tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
frmLingjian.LpXh.Close
frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
frmLingjian.dtgView.Columns(5).Caption = "伏斯价"
If mod1.VLP = 1 Then
    frmLingjian.dtgView.Columns("伏斯价").Visible = False
ElseIf mod1.VLP = 2 Then
    frmLingjian.dtgView.Columns("伏斯价").Visible = True
ElseIf mod1.VLP = 3 Then
    frmLingjian.dtgView.Columns("伏斯价").Visible = True
End If
    Set frmLingjian.dtgView.DataSource = Nothing
If mod1.VLP = 3 Then
    frmLingjian.cmdKq.Visible = True
Else
    frmLingjian.cmdKq.Visible = False
End If
cmdGx.Enabled = False
End Sub

Private Sub cmdMk_Click()
Set frmLingjian.LpXh = New ADODB.Recordset
Dim tt As String
Dim oo As Integer
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
End If
frmLingjian.Caption = "麦克威尔"
frmLingjian.Show

For oo = frmLingjian.comJzXh.ListCount - 1 To 0 Step -1
    frmLingjian.comJzXh.RemoveItem oo
Next

tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
frmLingjian.LpXh.Close
frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
frmLingjian.dtgView.Columns(5).Caption = "库存价"
If mod1.VLP = 1 Then
    frmLingjian.dtgView.Columns("库存价").Visible = False
ElseIf mod1.VLP = 2 Then
    frmLingjian.dtgView.Columns("库存价").Visible = True
ElseIf mod1.VLP = 3 Then
    frmLingjian.dtgView.Columns("库存价").Visible = True
End If
    Set frmLingjian.dtgView.DataSource = Nothing

If mod1.VLP = 3 Then
    frmLingjian.cmdKq.Visible = True
Else
    frmLingjian.cmdKq.Visible = False
End If
cmdGx.Enabled = False
End Sub

Private Sub cmdNew_Click()
frmGXBj.Show
frmGxBiao.Enabled = False
End Sub

Private Sub cmdPj_Click()
Dim pk As String
Set frmLingPei.LpXh = New ADODB.Recordset
Set frmLingPei.adoLpg = New ADODB.Recordset
Dim tt As String
Dim oo As Integer
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
End If

frmLingPei.Show


tt = "lpg_pei('')"
frmLingPei.adoLpg.Close
frmLingPei.adoLpg.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
'Set frmLingPei.mga.DataSource = frmLingPei.adoLpg
Set frmLingPei.mgb.DataSource = frmLingPei.adoLpg
 
'Set frmLingPei.mgc.DataSource = frmLingPei.adoLpg
pk = "<        |<      种  类          |<  品  牌     |<  型  号           |< 规  格     |< 面  价  |< 建议售价    |<   成本价   |<  进  价    "
frmLingPei.mgb.FormatString = pk
If mod1.VLP = 1 Then
       frmLingPei.mgb.ColWidth(8) = 0
ElseIf mod1.VLP = 2 Then
       frmLingPei.mgb.ColWidth(8) = -1
ElseIf mod1.VLP = 3 Then
       frmLingPei.mgb.ColWidth(8) = -1
End If
'    Set frmlingpei.dtgView.DataSource = Nothing
End Sub

Private Sub cmdTl_Click()
Set frmLingjian.LpXh = New ADODB.Recordset
Dim tt As String
Dim oo As Integer
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
End If
frmLingjian.Caption = "特灵"
frmLingjian.Show

For oo = frmLingjian.comJzXh.ListCount - 1 To 0 Step -1
    frmLingjian.comJzXh.RemoveItem oo
Next

tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
frmLingjian.LpXh.Close
frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
frmLingjian.dtgView.Columns(5).Caption = "库存价"
If mod1.VLP = 1 Then
    frmLingjian.dtgView.Columns("库存价").Visible = False
ElseIf mod1.VLP = 2 Then
    frmLingjian.dtgView.Columns("库存价").Visible = True
ElseIf mod1.VLP = 3 Then
    frmLingjian.dtgView.Columns("库存价").Visible = True
End If
    Set frmLingjian.dtgView.DataSource = Nothing
    
If mod1.VLP = 3 Then
    frmLingjian.cmdKq.Visible = True
Else
    frmLingjian.cmdKq.Visible = False
End If
cmdGx.Enabled = False
End Sub

Private Sub cmdYk_Click()
Set frmLingjian.LpXh = New ADODB.Recordset
Dim oo As Integer
Dim tt As String
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
End If
frmLingjian.Caption = "约克"
frmLingjian.Show
MsgBox "约克所有配件在2006年度均上涨10%以上，新价格暂未上传，报价及销售时请询问采购人员，谢谢!"
For oo = frmLingjian.comJzXh.ListCount - 1 To 0 Step -1
    frmLingjian.comJzXh.RemoveItem oo
Next
tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
frmLingjian.LpXh.Close
frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    frmLingjian.dtgView.Columns("伏斯价").Visible = False
        Set frmLingjian.dtgView.DataSource = Nothing
'If mod1.VLP = 1 Then
'    frmLingjian.dtgView.Columns("伏斯价").Visible = False
'ElseIf mod1.VLP = 2 Then
'    frmLingjian.dtgView.Columns("伏斯价").Visible = True
'ElseIf mod1.VLP = 3 Then
'    frmLingjian.dtgView.Columns("伏斯价").Visible = True
'End If
If mod1.VLP = 3 Then
    frmLingjian.cmdKq.Visible = True
Else
    frmLingjian.cmdKq.Visible = False
End If
cmdGx.Enabled = False
End Sub

Private Sub cmdZlG_Click()
MsgBox "注:以上价格有效期至2005年11月20日"
Set frmLingjian.LpXh = New ADODB.Recordset
Dim oo As Integer
Dim tt As String
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
End If
frmLingjian.Caption = "制冷剂"
frmLingjian.Show

For oo = frmLingjian.comJzXh.ListCount - 1 To 0 Step -1
    frmLingjian.comJzXh.RemoveItem oo
Next
tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
frmLingjian.LpXh.Close
frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    frmLingjian.dtgView.Columns("伏斯价").Visible = False
        Set frmLingjian.dtgView.DataSource = Nothing
        
If mod1.VLP = 3 Then
    frmLingjian.cmdKq.Visible = True
Else
    frmLingjian.cmdKq.Visible = False
End If
cmdGx.Enabled = False
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
frmGxBiao.Width = mod1.Fwidth
frmGxBiao.Height = mod1.FHeight
End Sub


Private Sub Form_Unload(Cancel As Integer)
Cancel = True
frmGxBiao.Visible = False
frmZu.Enabled = True
End Sub


