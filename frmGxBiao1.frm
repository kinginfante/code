VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGxBiao 
   Caption         =   "ѯ�ۼ�¼��"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdNew 
      Caption         =   "�½�ѯ�۵�"
      Height          =   735
      Left            =   12060
      TabIndex        =   10
      Top             =   5820
      Width           =   1725
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   555
      Left            =   14580
      Picture         =   "frmGxBiao1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8610
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������ݿ�"
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
         Caption         =   "Լ��"
         Height          =   585
         Left            =   4890
         TabIndex        =   6
         Top             =   1290
         Width           =   1815
      End
      Begin VB.CommandButton cmdTl 
         Caption         =   "����"
         Height          =   585
         Left            =   4890
         TabIndex        =   5
         Top             =   2580
         Width           =   1785
      End
      Begin VB.CommandButton cmdMk 
         Caption         =   "�������"
         Height          =   615
         Left            =   4890
         TabIndex        =   4
         Top             =   3180
         Width           =   1785
      End
      Begin VB.CommandButton cmdPj 
         Caption         =   "�������"
         Height          =   585
         Left            =   1890
         TabIndex        =   3
         Top             =   3210
         Width           =   1665
      End
      Begin VB.CommandButton cmdZlG 
         Caption         =   "�����"
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
'MsgBox "����!Ŀǰ�ٺ���ʲ�Ľ������в��죬�������޸�֮�У�����ĳɱ��۸�����԰���ǰ�ļ��㡣����Ʒ��û�б仯��лл  С�� �ֻ�111"
frmLingjian.Caption = "�ٺ���ʲ"
frmLingjian.Show

For oo = frmLingjian.comJzXh.ListCount - 1 To 0 Step -1
    frmLingjian.comJzXh.RemoveItem oo
Next

tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
frmLingjian.LpXh.Close
frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
frmLingjian.dtgView.Columns(5).Caption = "����"
If mod1.VLP = 1 Then
    frmLingjian.dtgView.Columns("����").Visible = False
ElseIf mod1.VLP = 2 Then
    frmLingjian.dtgView.Columns("����").Visible = True
ElseIf mod1.VLP = 3 Then
    frmLingjian.dtgView.Columns("����").Visible = True
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
frmLingjian.Caption = "����"
frmLingjian.Show

For oo = frmLingjian.comJzXh.ListCount - 1 To 0 Step -1
    frmLingjian.comJzXh.RemoveItem oo
Next

tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
frmLingjian.LpXh.Close
frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
frmLingjian.dtgView.Columns(5).Caption = "��˹��"
If mod1.VLP = 1 Then
    frmLingjian.dtgView.Columns("��˹��").Visible = False
ElseIf mod1.VLP = 2 Then
    frmLingjian.dtgView.Columns("��˹��").Visible = True
ElseIf mod1.VLP = 3 Then
    frmLingjian.dtgView.Columns("��˹��").Visible = True
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
frmLingjian.Caption = "�������"
frmLingjian.Show

For oo = frmLingjian.comJzXh.ListCount - 1 To 0 Step -1
    frmLingjian.comJzXh.RemoveItem oo
Next

tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
frmLingjian.LpXh.Close
frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
frmLingjian.dtgView.Columns(5).Caption = "����"
If mod1.VLP = 1 Then
    frmLingjian.dtgView.Columns("����").Visible = False
ElseIf mod1.VLP = 2 Then
    frmLingjian.dtgView.Columns("����").Visible = True
ElseIf mod1.VLP = 3 Then
    frmLingjian.dtgView.Columns("����").Visible = True
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
pk = "<        |<      ��  ��          |<  Ʒ  ��     |<  ��  ��           |< ��  ��     |< ��  ��  |< �����ۼ�    |<   �ɱ���   |<  ��  ��    "
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
frmLingjian.Caption = "����"
frmLingjian.Show

For oo = frmLingjian.comJzXh.ListCount - 1 To 0 Step -1
    frmLingjian.comJzXh.RemoveItem oo
Next

tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
frmLingjian.LpXh.Close
frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
frmLingjian.dtgView.Columns(5).Caption = "����"
If mod1.VLP = 1 Then
    frmLingjian.dtgView.Columns("����").Visible = False
ElseIf mod1.VLP = 2 Then
    frmLingjian.dtgView.Columns("����").Visible = True
ElseIf mod1.VLP = 3 Then
    frmLingjian.dtgView.Columns("����").Visible = True
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
frmLingjian.Caption = "Լ��"
frmLingjian.Show
MsgBox "Լ�����������2006��Ⱦ�����10%���ϣ��¼۸���δ�ϴ������ۼ�����ʱ��ѯ�ʲɹ���Ա��лл!"
For oo = frmLingjian.comJzXh.ListCount - 1 To 0 Step -1
    frmLingjian.comJzXh.RemoveItem oo
Next
tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
frmLingjian.LpXh.Close
frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    frmLingjian.dtgView.Columns("��˹��").Visible = False
        Set frmLingjian.dtgView.DataSource = Nothing
'If mod1.VLP = 1 Then
'    frmLingjian.dtgView.Columns("��˹��").Visible = False
'ElseIf mod1.VLP = 2 Then
'    frmLingjian.dtgView.Columns("��˹��").Visible = True
'ElseIf mod1.VLP = 3 Then
'    frmLingjian.dtgView.Columns("��˹��").Visible = True
'End If
If mod1.VLP = 3 Then
    frmLingjian.cmdKq.Visible = True
Else
    frmLingjian.cmdKq.Visible = False
End If
cmdGx.Enabled = False
End Sub

Private Sub cmdZlG_Click()
MsgBox "ע:���ϼ۸���Ч����2005��11��20��"
Set frmLingjian.LpXh = New ADODB.Recordset
Dim oo As Integer
Dim tt As String
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
End If
frmLingjian.Caption = "�����"
frmLingjian.Show

For oo = frmLingjian.comJzXh.ListCount - 1 To 0 Step -1
    frmLingjian.comJzXh.RemoveItem oo
Next
tt = "LPG_jzXhP('" & frmLingjian.Caption & "')"
frmLingjian.LpXh.Close
frmLingjian.LpXh.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    frmLingjian.dtgView.Columns("��˹��").Visible = False
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


