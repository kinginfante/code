VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGxBiao 
   BackColor       =   &H00FFFFC0&
   Caption         =   "ѯ�ۼ�¼��"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdW 
      Caption         =   "��������"
      Height          =   345
      Left            =   8250
      TabIndex        =   28
      Top             =   8670
      Width           =   1335
   End
   Begin VB.CommandButton cmdQH 
      Caption         =   "��ʾ����"
      Height          =   1005
      Left            =   12540
      TabIndex        =   26
      Top             =   6930
      Width           =   315
   End
   Begin VB.CommandButton cmdZF 
      Caption         =   "����"
      Height          =   645
      Left            =   12540
      TabIndex        =   25
      Top             =   6180
      Width           =   315
   End
   Begin VB.Frame frmC 
      Caption         =   "��ѯ"
      Height          =   705
      Left            =   30
      TabIndex        =   18
      Top             =   8400
      Visible         =   0   'False
      Width           =   8115
      Begin VB.CommandButton cmdAll 
         Caption         =   "ȫ��"
         Height          =   285
         Left            =   7140
         TabIndex        =   24
         Top             =   300
         Width           =   885
      End
      Begin VB.ComboBox comLx 
         Height          =   300
         ItemData        =   "frmGxBiao.frx":0000
         Left            =   870
         List            =   "frmGxBiao.frx":0016
         TabIndex        =   21
         Text            =   "��Ʒ����"
         Top             =   300
         Width           =   1965
      End
      Begin VB.TextBox txtZ 
         Height          =   285
         Left            =   3840
         TabIndex        =   20
         Top             =   300
         Width           =   1875
      End
      Begin VB.CommandButton cmdC 
         Caption         =   "��ѯ"
         Height          =   285
         Left            =   5880
         TabIndex        =   19
         Top             =   330
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "��ѯ����"
         Height          =   225
         Left            =   0
         TabIndex        =   23
         Top             =   330
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "ֵ"
         Height          =   255
         Left            =   2940
         TabIndex        =   22
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "�� ��"
      Height          =   405
      Left            =   12960
      TabIndex        =   6
      Top             =   90
      Width           =   2265
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   555
      Left            =   14580
      Picture         =   "frmGxBiao.frx":0050
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8610
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������ݿ�"
      Height          =   5025
      Left            =   13170
      TabIndex        =   0
      Top             =   3030
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton cmdPj 
         Caption         =   "�������"
         Height          =   315
         Left            =   300
         TabIndex        =   9
         Top             =   1220
         Width           =   1365
      End
      Begin VB.CommandButton cmdMk 
         Caption         =   "�������"
         Height          =   315
         Left            =   300
         TabIndex        =   8
         Top             =   3720
         Width           =   1365
      End
      Begin VB.CommandButton cmdYk 
         Caption         =   "Լ��"
         Height          =   315
         Left            =   300
         TabIndex        =   7
         Top             =   2720
         Width           =   1365
      End
      Begin VB.CommandButton cmdDunham 
         Caption         =   "�ٺ���ʲ"
         Height          =   315
         Left            =   300
         Picture         =   "frmGxBiao.frx":0152
         TabIndex        =   4
         Top             =   1720
         Width           =   1365
      End
      Begin VB.CommandButton cmdKl 
         BackColor       =   &H80000009&
         Caption         =   "����"
         Height          =   315
         Left            =   300
         Picture         =   "frmGxBiao.frx":0E2B
         TabIndex        =   3
         Top             =   2220
         Width           =   1365
      End
      Begin VB.CommandButton cmdTl 
         Caption         =   "����"
         Height          =   315
         Left            =   300
         TabIndex        =   2
         Top             =   3220
         Width           =   1365
      End
      Begin VB.CommandButton cmdZlG 
         Caption         =   "�����"
         Height          =   315
         Left            =   300
         TabIndex        =   1
         Top             =   720
         Width           =   1365
      End
   End
   Begin VB.Frame frmNew 
      BorderStyle     =   0  'None
      Height          =   4965
      Left            =   8820
      TabIndex        =   10
      Top             =   9120
      Visible         =   0   'False
      Width           =   5505
      Begin VB.CommandButton cmdCreat 
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1140
         Width           =   1245
      End
      Begin VB.CommandButton cmdDx 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2700
         Width           =   1245
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "ά��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3480
         Width           =   1245
      End
      Begin VB.CommandButton cmdFb 
         Caption         =   "���̷ְ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   4260
         Width           =   1245
      End
      Begin VB.CommandButton cmdCP 
         Caption         =   "��Ʒ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "�½�ѯ�۵�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   17
         Top             =   450
         Width           =   1425
      End
      Begin VB.Shape Shape1 
         Height          =   3975
         Left            =   1830
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   3345
      End
      Begin VB.Label lblZM 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3135
         Left            =   2190
         TabIndex        =   16
         Top             =   1230
         Width           =   2625
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgXj 
      Height          =   7965
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   12465
      _ExtentX        =   21987
      _ExtentY        =   14049
      _Version        =   393216
      BackColor       =   16777152
      BackColorFixed  =   12648384
      BackColorBkg    =   16777152
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmGxBiao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoXj As Object
Public adoGc As Object


Private Sub cmdAll_Click()
Dim tt As String
On Error Resume Next


If mod1.Bm = "�����ҵ��" Or mod1.DName = "�ܴ���" Then

    tt = "select * from xunjiaQ where pz='�����' order by ѯ������ desc"

ElseIf mod1.Bm = "���̲�" Or mod1.Bm = "���ݹ��̲�" Then
     If mod1.DName = "����" Then
        tt = "select * from xunjiaView where qy='�Ϻ�' and pz<>'��Ʒ' and comid=0 and lc>=4"

'    ElseIf mod1.DName = "֣��" Then
'        tt = "select * from xunjiaView where qy<>'�Ϻ�' and pz<>'��Ʒ' and comid=0 and lc>=4"
'    ElseIf mod1.DName = "����" Then
'        tt = "select * from xunjiaView where comid=1 and pz<>'��Ʒ' and lc>=4"
    Else '�鳤
        tt = "select * from xunJiaView where  zh=" & mod1.Gzu & "  and lc>=3"
    End If
End If

frmGxBiao.adoXj.Close
frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
If frmGxBiao.adoXj.RecordCount > 1 Then
    frmGxBiao.dtgXj.FixedRows = 0
    frmGxBiao.dtgXj.MergeCol(1) = True
    frmGxBiao.dtgXj.MergeCol(2) = True
    frmGxBiao.dtgXj.MergeCol(3) = True
    frmGxBiao.dtgXj.MergeCol(4) = True
    frmGxBiao.dtgXj.MergeCol(5) = True
    frmGxBiao.dtgXj.MergeCells = 3
    frmGxBiao.dtgXj.FixedRows = 1
End If
End Sub

Private Sub cmdBack_Click()
frmGxBiao.Visible = False
frmZu.Enabled = True
End Sub

Private Sub cmdC_Click()
Dim tt As String
On Error Resume Next

If mod1.Bm = "�����ҵ��" Or mod1.DName = "�ܴ���" Or mod1.DName = "" Or mod1.DName = "�����" Or mod1.DName = "" Then
    Select Case comLx.Text
    Case "��Ʒ����"
        tt = "select * from xunjiaQ where   ��Ʒ���� like '%" & txtZ.Text & "%' order by ѯ������ desc"
    Case "����ͺ�"
        tt = "select * from xunjiaQ where   ����ͺ� like '%" & txtZ.Text & "%' order by ѯ������ desc"
    Case "����"
        tt = "select * from xunjiaQ where  year(ѯ������)=" & Year(txtZ.Text) & " and month(ѯ������)=" & Month(txtZ.Text) & " and day(ѯ������)=" & _
        Day(txtZ.Text)
    Case "ҵ��Ա"
        tt = "select * from xunjiaQ where   ywy='" & txtZ.Text & "' order by ѯ������ desc"
    Case "��Ŀ����"
        tt = "select * from xunjiaQ where   ��Ŀ���� like '%" & txtZ.Text & "%' order by ѯ������ desc"
    Case "����Ʒ��"
        tt = "select * from xunjiaQ where   ����Ʒ�� like '%" & txtZ.Text & "%' order by ѯ������ desc"
    End Select
ElseIf mod1.Bm = "���̲�" Or mod1.Bm = "���ݹ��̲�" Or Mid(mod1.Bm, 3, 2) = "����" Then
     If mod1.DName = "����" Or mod1.DName = "����" Or mod1.DName = "�����" Or mod1.DName = "֣��" Or mod1.DName = "���h" Then
        tt = "select * from xunjiaView where  ����<>'����' and  ��Ŀ���� like '%" & txtZ.Text & "%' and comid=0 and lc>=4"

'    ElseIf mod1.DName = "֣��" Then
'        tt = "select * from xunjiaView where qy<>'�Ϻ�' and ����<>'����' and  ��Ŀ���� like '%" & txtZ.Text & "%'  and comid=0 and lc>=4"
'    ElseIf mod1.DName = "����" Then
'        tt = "select * from xunjiaView where comid=1 and ����<>'����' and  ��Ŀ���� like '%" & txtZ.Text & "%'  and lc>=4"
'''''    Else '�鳤
'''''        tt = "select * from xunJiaView where ����<>'����' and zh=" & mod1.Gzu & " and  ��Ŀ���� like '%" & txtZ.Text & "%'   and lc>=3"
'''''    End If
    End If
End If

frmGxBiao.adoXj.Close
frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
If frmGxBiao.adoXj.RecordCount > 1 Then
    frmGxBiao.dtgXj.FixedRows = 0
    frmGxBiao.dtgXj.MergeCol(1) = True
    frmGxBiao.dtgXj.MergeCol(2) = True
    frmGxBiao.dtgXj.MergeCol(3) = True
    frmGxBiao.dtgXj.MergeCol(4) = True
    frmGxBiao.dtgXj.MergeCol(5) = True
    frmGxBiao.dtgXj.MergeCells = 3
    frmGxBiao.dtgXj.FixedRows = 1
End If
        



End Sub

Private Sub cmdCP_Click()
Dim tt As String
On Error Resume Next

mod1.BTZ = 36
frmGXBj.Visible = False
Call modBJD.BJDGXQing
Call modBJD.gxbjUnLocked
frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "xunJiaAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@ywy") = mod1.DName
mod1.cmd.Parameters("@uid") = mod1.DHid
mod1.cmd.Parameters("@Lx") = 0
mod1.cmd.Parameters("@zl") = "����"
mod1.cmd.Parameters("@Lcou") = Right(frmGxBiao.cmdCreat.ToolTipText, 1) '��������
mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
mod1.cmd.Parameters("@lcRen") = mod1.DName
mod1.cmd.Parameters("@lcUid") = mod1.DHid
mod1.cmd.Parameters("@nLb") = frmGxBiao.cmdCreat.Tag
mod1.cmd.Execute
frmGXBj.lblBid.Caption = mod1.cmd.Parameters("@bid").Value
frmGXBj.lblBh.Caption = "XJD" & mod1.cmd.Parameters("@bid").Value
frmGXBj.lblOid.Caption = mod1.cmd.Parameters("@bid").Value
frmGXBj.lblLcou.Caption = Right(frmGxBiao.cmdNew.ToolTipText, 1) '��������
frmGXBj.lblLc.Caption = 0
frmGXBj.lblLcRen.Caption = mod1.DName
frmGXBj.lblLcUid.Caption = mod1.DHid
frmGXBj.lblNlb.Caption = frmGxBiao.cmdCreat.Tag
frmGXBj.lblYwy.Caption = mod1.DName
frmGXBj.lblUid.Caption = mod1.DHid
frmGXBj.lblZl.Caption = "����"
frmGXBj.comLx.Text = "��Ʒ"
Set cmd = Nothing
If frmGXBj.lblBh.Caption = "" Then
    ii = MsgBox("ϵͳ������������,�����̹ر�!�ٴδ򿪺�����Ϣ,������˴���.", vbOKOnly + vbExclamation, "A������")
    End
End If
'������Ŀ������Ϣ
tt = "select xmmc,xid from xmzl where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"
frmGXBj.adoXm.Close
frmGXBj.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGXBj.comXmmc.RowSource = frmGXBj.adoXm
frmGXBj.comXmmc.ListField = "xmmc"
frmGXBj.comXmmc.BoundColumn = "xid"

tt = "select jzpb,pbid from bjxt_jzpb"
frmGXBj.adoPb.Close
frmGXBj.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGXBj.comJzpb.RowSource = frmGXBj.adoPb
frmGXBj.comJzpb.ListField = "jzpb"
frmGXBj.comJzpb.BoundColumn = "pbid"
frmGXBj.txtHg.Locked = True
frmGXBj.txtYhg.Locked = True

    '�������̰�ť
    Call modBJD.XJGXLcBut(43)
    
frmWait.Visible = False
frmGXBj.Visible = True
frmGXBj.cmdMod.Enabled = False
frmGXBj.frmCg.Enabled = False
'ˢ�¹����б�
tt = "select * from xunJIamxView where bid=" & Val(frmGXBj.lblBid.Caption)
    frmGXBj.adoGx.Close
    frmGXBj.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGXBj.dtgMa.DataSource = frmGXBj.adoGx

frmGXBj.cmdSave.Enabled = True
frmGxBiao.Enabled = False
'frmGXBj.cmdBjd.Visible = False
frmGXBj.txtYhg.Locked = True
frmGXBj.comXmmc.Locked = False
frmGXBj.lblZl.ForeColor = &HC000C0
frmGXBj.lblzlZ.ForeColor = &HC000C0
frmGXBj.txtMj.Locked = True
frmGXBj.txtDj.Locked = True
frmGXBj.comLx.ToolTipText = "��Ʒ"
End Sub

Private Sub cmdCP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblZM.Caption = "    ָ��˾��ͻ�ǩ���Ĳ�Ʒ��Ӧ��ͬ�����԰����ɲ�Ʒ��Ӧ���ṩ�ĵ��ԡ���װ�˹��ѣ���ѯ���ɺ�����˾ʵʩ�����漰������˾���̲����˹����á�"
End Sub


Private Sub cmdCreat_Click()

Dim tt As String
On Error Resume Next

mod1.BTZ = 36
frmGXBj.Visible = False
Call modBJD.BJDGXQing
Call modBJD.gxbjUnLocked
frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "xunJiaAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@ywy") = mod1.DName
mod1.cmd.Parameters("@uid") = mod1.DHid
mod1.cmd.Parameters("@Lx") = 0
mod1.cmd.Parameters("@zl") = "����"
mod1.cmd.Parameters("@Lcou") = Right(frmGxBiao.cmdCreat.ToolTipText, 1) '��������
mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
mod1.cmd.Parameters("@lcRen") = mod1.DName
mod1.cmd.Parameters("@lcUid") = mod1.DHid
mod1.cmd.Parameters("@nLb") = frmGxBiao.cmdCreat.Tag
mod1.cmd.Execute
frmGXBj.lblBid.Caption = mod1.cmd.Parameters("@bid").Value
frmGXBj.lblBh.Caption = "XJD" & mod1.cmd.Parameters("@bid").Value
frmGXBj.lblOid.Caption = mod1.cmd.Parameters("@bid").Value
frmGXBj.lblLcou.Caption = Right(frmGxBiao.cmdNew.ToolTipText, 1) '��������
frmGXBj.lblLc.Caption = 0
frmGXBj.lblLcRen.Caption = mod1.DName
frmGXBj.lblLcUid.Caption = mod1.DHid
frmGXBj.lblNlb.Caption = frmGxBiao.cmdCreat.Tag
frmGXBj.lblYwy.Caption = mod1.DName
frmGXBj.lblUid.Caption = mod1.DHid
frmGXBj.lblZl.Caption = "����"
frmGXBj.comLx.Text = "�����"
Set cmd = Nothing
If frmGXBj.lblBh.Caption = "" Then
    ii = MsgBox("ϵͳ������������,�����̹ر�!�ٴδ򿪺�����Ϣ,������˴���.", vbOKOnly + vbExclamation, "A������")
    End
End If
'������Ŀ������Ϣ
tt = "select xmmc,xid from xmzl where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"
frmGXBj.adoXm.Close
frmGXBj.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGXBj.comXmmc.RowSource = frmGXBj.adoXm
frmGXBj.comXmmc.ListField = "xmmc"
frmGXBj.comXmmc.BoundColumn = "xid"

tt = "select jzpb,pbid from bjxt_jzpb"
frmGXBj.adoPb.Close
frmGXBj.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGXBj.comJzpb.RowSource = frmGXBj.adoPb
frmGXBj.comJzpb.ListField = "jzpb"
frmGXBj.comJzpb.BoundColumn = "pbid"
frmGXBj.txtHg.Locked = True
frmGXBj.txtYhg.Locked = True

    '�������̰�ť
    Call modBJD.XJGXLcBut(43)
    
frmWait.Visible = False
frmGXBj.Visible = True
frmGXBj.cmdMod.Enabled = False
frmGXBj.frmCg.Enabled = False
'ˢ�¹����б�
tt = "select * from xunJIamxView where bid=" & Val(frmGXBj.lblBid.Caption)
    frmGXBj.adoGx.Close
    frmGXBj.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGXBj.dtgMa.DataSource = frmGXBj.adoGx

frmGXBj.cmdSave.Enabled = True
frmGxBiao.Enabled = False
'frmGXBj.cmdBjd.Visible = False
frmGXBj.txtYhg.Locked = True
frmGXBj.comXmmc.Locked = False
frmGXBj.lblZl.ForeColor = &HC000C0
frmGXBj.lblzlZ.ForeColor = &HC000C0
frmGXBj.txtMj.Locked = True
frmGXBj.txtDj.Locked = True
frmGXBj.comLx.ToolTipText = "�����"
End Sub

Private Sub cmdCreat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblZM.Caption = "    ָ��˾��ͻ�ǩ�����������Ӧ��ͬ�����漰������˾���̲����˹����ã�ѯ�۱���ͨ������ָ�����������˾��"
End Sub


Private Sub cmdDunham_Click()
Set frmLingjian.LpXh = CreateObject("adodb.recordset")
Dim tt As String
Dim oo As Integer
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
    Exit Sub
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

Private Sub cmdDx_Click()
Dim tt As String
On Error Resume Next
mod1.BTZ = 36
'���½�ά��ѯ��
frmWBXJ.Visible = False
Call modBJD.BJDWBQing
Call modBJD.wbxjUnLocked
frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "xunJiaAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@ywy") = mod1.DName
mod1.cmd.Parameters("@uid") = mod1.DHid
mod1.cmd.Parameters("@Lx") = 1
mod1.cmd.Parameters("@zl") = "����"
mod1.cmd.Parameters("@Lcou") = Right(frmGxBiao.cmdNew.ToolTipText, 1) '��������
mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
mod1.cmd.Parameters("@lcRen") = mod1.DName
mod1.cmd.Parameters("@lcUid") = mod1.DHid
mod1.cmd.Parameters("@nLb") = frmGxBiao.cmdNew.Tag
mod1.cmd.Execute
frmWBXJ.lblBid.Caption = mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblBh.Caption = "XJD" & mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblOid.Caption = mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblLcou.Caption = Right(frmGxBiao.cmdNew.ToolTipText, 1) '��������
frmWBXJ.lblLc.Caption = 0
frmWBXJ.lblLcRen.Caption = mod1.DName
frmWBXJ.lblLcUid.Caption = mod1.DHid
frmWBXJ.lblNlb.Caption = frmGxBiao.cmdNew.Tag
frmWBXJ.lblYwy.Caption = mod1.DName
frmWBXJ.lblUid.Caption = mod1.DHid
frmWBXJ.lblZl.Caption = "����"
Set cmd = Nothing
If frmWBXJ.lblBh.Caption = "" Then
    ii = MsgBox("ϵͳ������������,�����̹ر�!�ٴδ򿪺�����Ϣ,������˴���.", vbOKOnly + vbExclamation, "A������")
    End
End If
'������Ŀ������Ϣ
tt = "select xmmc,xid from xmzl where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"
frmWBXJ.adoXm.Close
frmWBXJ.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.comXmmc.RowSource = frmWBXJ.adoXm
frmWBXJ.comXmmc.ListField = "xmmc"
frmWBXJ.comXmmc.BoundColumn = "xid"

tt = "select jzpb,pbid from bjxt_jzpb"
frmWBXJ.adoPb.Close
frmWBXJ.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.comPb.RowSource = frmWBXJ.adoPb
frmWBXJ.comPb.ListField = "jzpb"
frmWBXJ.comPb.BoundColumn = "pbid"
            frmWBXJ.frmDx.Visible = True
            frmWBXJ.frmNb.Visible = False
            frmWBXJ.frmTime.Visible = False

            frmWBXJ.cmdD.Visible = False
            frmWBXJ.cmdJi.Visible = False
            frmWBXJ.tabGc.TabVisible(2) = True
            frmWBXJ.tabGc.TabVisible(0) = False
            frmWBXJ.tabGc.TabVisible(1) = False
            frmWBXJ.tabGc.Tab = 2

    '�������̰�ť
    Call modBJD.XJWBLcBut(44)
    
'������Ϣ��
frmWBXJ.frmNew.Visible = True
tt = "select jzpb as ����Ʒ��,jzxh as �����ͺ�,sl as ����,jxId from wbjb where bid=" & Val(frmWBXJ.lblBid.Caption)
Set frmWBXJ.adoA = CreateObject("adodb.recordset")
frmWBXJ.adoA.Close
frmWBXJ.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.dtgA.DataSource = frmWBXJ.adoA
frmWBXJ.cmdTk.Visible = False
    
frmWait.Visible = False
frmWBXJ.Visible = True
frmWBXJ.cmdMod.Enabled = False
frmWBXJ.txtMOn.Locked = False
frmWBXJ.txtHg.Locked = True
frmWBXJ.txtYhg.Locked = True
frmWBXJ.txtClf.Locked = True
frmWBXJ.cmdSave.Enabled = True
End Sub

Private Sub cmdDx_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblZM.Caption = "    ָ��˾��ͻ�ǩ���ġ�����˾�Ĺ�����Ա���豸���е�һ����ά�޺�ͬ�����Ժ����������Ʒ������ҵ��ְ�������ά�������ı�֤�ڲ�����6���¡�"
End Sub


Private Sub cmdFb_Click()
'Dim tt As String
'On Error Resume Next
'
'mod1.BTZ = 36
'frmGXBj.Visible = False
'Call modBJD.BJDGXQing
'Call modBJD.gxbjUnLocked
'frmWait.Show
'frmWait.ZOrder 0
'frmWait.Refresh
'Set mod1.cmd = createobject("adodb.command")
'mod1.cmd.ActiveConnection = mod1.CC
'mod1.cmd.CommandText = "xunJiaAdd"
'mod1.cmd.CommandType = adCmdStoredProc
'mod1.cmd.Parameters("@ywy") = mod1.DName
'mod1.cmd.Parameters("@uid") = mod1.DHid
'mod1.cmd.Parameters("@Lx") = 0
'mod1.cmd.Parameters("@zl") = "����"
'mod1.cmd.Parameters("@Lcou") = Right(frmGxBiao.cmdCreat.ToolTipText, 1) '��������
'mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
'mod1.cmd.Parameters("@lcRen") = mod1.DName
'mod1.cmd.Parameters("@lcUid") = mod1.DHid
'mod1.cmd.Parameters("@nLb") = frmGxBiao.cmdCreat.Tag
'mod1.cmd.Execute
'frmGXBj.lblBid.Caption = mod1.cmd.Parameters("@bid").Value
'frmGXBj.lblBh.Caption = "XJD" & mod1.cmd.Parameters("@bid").Value
'frmGXBj.lblOid.Caption = mod1.cmd.Parameters("@bid").Value
'frmGXBj.lblLcou.Caption = Right(frmGxBiao.cmdNew.ToolTipText, 1) '��������
'frmGXBj.lblLc.Caption = 0
'frmGXBj.lblLcRen.Caption = mod1.DName
'frmGXBj.lblLcUid.Caption = mod1.DHid
'frmGXBj.lblNlb.Caption = frmGxBiao.cmdCreat.Tag
'frmGXBj.lblYwy.Caption = mod1.DName
'frmGXBj.lblUid.Caption = mod1.DHid
'frmGXBj.lblZl.Caption = "����"
'frmGXBj.comLx.Text = "��Ʒ"
'Set cmd = Nothing
'If frmGXBj.lblBh.Caption = "" Then
'    ii = MsgBox("ϵͳ������������,�����̹ر�!�ٴδ򿪺�����Ϣ,������˴���.", vbOKOnly + vbExclamation, "A������")
'    End
'End If
''������Ŀ������Ϣ
'tt = "select xmmc,xid from xmzl where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"
'frmGXBj.adoXm.Close
'frmGXBj.adoXm.Open tt, mod1.workkk, adOpenKeyset, adLockReadOnly, adCmdText
'Set frmGXBj.comXmmc.RowSource = frmGXBj.adoXm
'frmGXBj.comXmmc.ListField = "xmmc"
'frmGXBj.comXmmc.BoundColumn = "xid"
'
'tt = "select jzpb,pbid from bjxt_jzpb"
'frmGXBj.adoPb.Close
'frmGXBj.adoPb.Open tt, mod1.workkk, adOpenKeyset, adLockReadOnly, adCmdText
'Set frmGXBj.comJzpb.RowSource = frmGXBj.adoPb
'frmGXBj.comJzpb.ListField = "jzpb"
'frmGXBj.comJzpb.BoundColumn = "pbid"
'frmGXBj.txtHg.Locked = True
'frmGXBj.txtYhg.Locked = True
'
'    '�������̰�ť
'    Call modBJD.XJGXLcBut(43)
'
'frmWait.Visible = False
'frmGXBj.Visible = True
'frmGXBj.cmdMod.Enabled = False
'frmGXBj.frmCg.Enabled = False
''ˢ�¹����б�
'tt = "select * from xunJIamxView where bid=" & Val(frmGXBj.lblBid.Caption)
'    frmGXBj.adoGx.Close
'    frmGXBj.adoGx.Open tt, mod1.workkk, adOpenKeyset, adLockReadOnly, adCmdText
'    Set frmGXBj.dtgMa.DataSource = frmGXBj.adoGx
'
'frmGXBj.cmdSave.Enabled = True
'frmGxBiao.Enabled = False
''frmGXBj.cmdBjd.Visible = False
'frmGXBj.txtYhg.Locked = True
'frmGXBj.comXmmc.Locked = False
'frmGXBj.lblZl.ForeColor = &HC000C0
'frmGXBj.lblzlZ.ForeColor = &HC000C0
'frmGXBj.txtMj.Locked = True
'frmGXBj.txtDj.Locked = True
'frmGXBj.FB = True

Dim tt As String
On Error Resume Next
mod1.BTZ = 36
'���½�ά��ѯ��
frmWBXJ.Visible = False
Call modBJD.BJDWBQing
Call modBJD.wbxjUnLocked
frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "xunJiaAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@ywy") = mod1.DName
mod1.cmd.Parameters("@uid") = mod1.DHid
mod1.cmd.Parameters("@Lx") = 1
mod1.cmd.Parameters("@zl") = "���̷ְ�"
mod1.cmd.Parameters("@Lcou") = Right(frmGxBiao.cmdNew.ToolTipText, 1) '��������
mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
mod1.cmd.Parameters("@lcRen") = mod1.DName
mod1.cmd.Parameters("@lcUid") = mod1.DHid
mod1.cmd.Parameters("@nLb") = frmGxBiao.cmdNew.Tag
mod1.cmd.Execute
frmWBXJ.lblBid.Caption = mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblBh.Caption = "XJD" & mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblOid.Caption = mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblLcou.Caption = Right(frmGxBiao.cmdNew.ToolTipText, 1) '��������
frmWBXJ.lblLc.Caption = 0
frmWBXJ.lblLcRen.Caption = mod1.DName
frmWBXJ.lblLcUid.Caption = mod1.DHid
frmWBXJ.lblNlb.Caption = frmGxBiao.cmdNew.Tag
frmWBXJ.lblYwy.Caption = mod1.DName
frmWBXJ.lblUid.Caption = mod1.DHid
frmWBXJ.lblZl.Caption = "���̷ְ�"
Set cmd = Nothing
If frmWBXJ.lblBh.Caption = "" Then
    ii = MsgBox("ϵͳ������������,�����̹ر�!�ٴδ򿪺�����Ϣ,������˴���.", vbOKOnly + vbExclamation, "A������")
    End
End If
'������Ŀ������Ϣ
tt = "select xmmc,xid from xmzl where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"
frmWBXJ.adoXm.Close
frmWBXJ.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.comXmmc.RowSource = frmWBXJ.adoXm
frmWBXJ.comXmmc.ListField = "xmmc"
frmWBXJ.comXmmc.BoundColumn = "xid"

tt = "select jzpb,pbid from bjxt_jzpb"
frmWBXJ.adoPb.Close
frmWBXJ.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.comPb.RowSource = frmWBXJ.adoPb
frmWBXJ.comPb.ListField = "jzpb"
frmWBXJ.comPb.BoundColumn = "pbid"
            frmWBXJ.frmDx.Visible = True
            frmWBXJ.frmNb.Visible = False
            frmWBXJ.frmTime.Visible = False

            frmWBXJ.cmdD.Visible = False
            frmWBXJ.cmdJi.Visible = False
            frmWBXJ.tabGc.TabVisible(2) = True
            frmWBXJ.tabGc.TabVisible(0) = False
            frmWBXJ.tabGc.TabVisible(1) = False
            frmWBXJ.tabGc.Tab = 2

    '�������̰�ť
    Call modBJD.XJWBLcBut(44)
    
'������Ϣ��
frmWBXJ.frmNew.Visible = True
tt = "select jzpb as ����Ʒ��,jzxh as �����ͺ�,sl as ����,jxId from wbjb where bid=" & Val(frmWBXJ.lblBid.Caption)
Set frmWBXJ.adoA = CreateObject("adodb.recordset")
frmWBXJ.adoA.Close
frmWBXJ.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.dtgA.DataSource = frmWBXJ.adoA
frmWBXJ.cmdTk.Visible = False
    
frmWait.Visible = False
frmWBXJ.Visible = True
frmWBXJ.cmdMod.Enabled = False
frmWBXJ.txtMOn.Locked = False
frmWBXJ.txtHg.Locked = True
frmWBXJ.txtYhg.Locked = True
frmWBXJ.txtClf.Locked = True
frmWBXJ.cmdSave.Enabled = True
frmWBXJ.Caption = frmWBXJ.Caption & "(���̷ְ�)"
End Sub

Private Sub cmdFb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblZM.Caption = "    ָ��˾��ͻ�ǩ���İ�װ��ά�ޡ������ͬ�����в��漰������˾���̲����˹����ã��˹�����ȫ���ɷְ��̳е�����������Ʒ������������ۡ�"
End Sub

Private Sub cmdGC_Click()
Dim tt As String
On Error Resume Next

mod1.BTZ = 36
frmGXBj.Visible = False
Call modBJD.BJDGXQing
Call modBJD.gxbjUnLocked
frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "xunJiaAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@ywy") = mod1.DName
mod1.cmd.Parameters("@uid") = mod1.DHid
mod1.cmd.Parameters("@Lx") = 0
mod1.cmd.Parameters("@zl") = "����"
mod1.cmd.Parameters("@Lcou") = 3 '��������
mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
mod1.cmd.Parameters("@lcRen") = mod1.DName
mod1.cmd.Parameters("@lcUid") = mod1.DHid
mod1.cmd.Parameters("@nLb") = 43
mod1.cmd.Execute
frmGXBj.lblBid.Caption = mod1.cmd.Parameters("@bid").Value
frmGXBj.lblBh.Caption = "XJD" & mod1.cmd.Parameters("@bid").Value
frmGXBj.lblOid.Caption = mod1.cmd.Parameters("@bid").Value
frmGXBj.lblLcou.Caption = 3 '��������
frmGXBj.lblLc.Caption = 0
frmGXBj.lblLcRen.Caption = mod1.DName
frmGXBj.lblLcUid.Caption = mod1.DHid
frmGXBj.lblNlb.Caption = 43
frmGXBj.lblYwy.Caption = mod1.DName
frmGXBj.lblUid.Caption = mod1.DHid
frmGXBj.lblZl.Caption = "����"
frmGXBj.comLx.Text = "�����"
Set cmd = Nothing
If frmGXBj.lblBh.Caption = "" Then
    ii = MsgBox("ϵͳ������������,�����̹ر�!�ٴδ򿪺�����Ϣ,������˴���.", vbOKOnly + vbExclamation, "A������")
    End
End If
'������Ŀ������Ϣ
'tt = "select xmmc,xid from xmzl where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"
tt = "select xmmc,xid from wbZname where zname='" & mod1.DName & "' and uid='" & mod1.DHid & "' order by xmmc"
frmGXBj.adoXm.Close
frmGXBj.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGXBj.comXmmc.RowSource = frmGXBj.adoXm
frmGXBj.comXmmc.ListField = "xmmc"
frmGXBj.comXmmc.BoundColumn = "xid"

tt = "select jzpb,pbid from bjxt_jzpb"
frmGXBj.adoPb.Close
frmGXBj.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGXBj.comJzpb.RowSource = frmGXBj.adoPb
frmGXBj.comJzpb.ListField = "jzpb"
frmGXBj.comJzpb.BoundColumn = "pbid"
frmGXBj.txtHg.Locked = True
frmGXBj.txtYhg.Locked = True

    '�������̰�ť
    Call modBJD.XJGXLcBut(43)
    
frmWait.Visible = False
frmGXBj.Visible = True
frmGXBj.cmdMod.Enabled = False
frmGXBj.frmCg.Enabled = False
'ˢ�¹����б�
tt = "select * from xunJIamxView where bid=" & Val(frmGXBj.lblBid.Caption)
    frmGXBj.adoGx.Close
    frmGXBj.adoGx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGXBj.dtgMa.DataSource = frmGXBj.adoGx

frmGXBj.cmdSave.Enabled = True
frmGxBiao.Enabled = False
'frmGXBj.cmdBjd.Visible = False
frmGXBj.txtYhg.Locked = True
frmGXBj.comXmmc.Locked = False
frmGXBj.lblZl.ForeColor = &HC000C0
frmGXBj.lblzlZ.ForeColor = &HC000C0
frmGXBj.txtMj.Locked = True
frmGXBj.txtDj.Locked = True
frmGXBj.comLx.ToolTipText = "�����"
End Sub

Private Sub cmdKl_Click()
Set frmLingjian.LpXh = CreateObject("adodb.recordset")
Dim tt As String
Dim oo As Integer
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
    Exit Sub
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
    frmLingjian.dtgView.Columns("��˹��").Visible = False
ElseIf mod1.VLP = 3 Then
    frmLingjian.dtgView.Columns("��˹��").Visible = True
End If
    Set frmLingjian.dtgView.DataSource = Nothing
If mod1.DName = "�Ŵ���" Then
    frmLingjian.cmdKq.Visible = True
Else
    frmLingjian.cmdKq.Visible = False
End If
cmdGx.Enabled = False
End Sub

Private Sub cmdMk_Click()
Set frmLingjian.LpXh = CreateObject("adodb.recordset")
Dim tt As String
Dim oo As Integer
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
    Exit Sub
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
    frmLingjian.frmMod.Visible = True
Else
    frmLingjian.cmdKq.Visible = False
    frmLingjian.frmMod.Visible = False
End If
cmdGx.Enabled = False
End Sub

Private Sub cmdNew_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
mod1.BTZ = 36
'���½�ά��ѯ��
frmWBXJ.Visible = False
Call modBJD.BJDWBQing
Call modBJD.wbxjUnLocked
frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "xunJiaAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@ywy") = mod1.DName
mod1.cmd.Parameters("@uid") = mod1.DHid
mod1.cmd.Parameters("@Lx") = 1
mod1.cmd.Parameters("@zl") = "ά��"
mod1.cmd.Parameters("@Lcou") = Right(frmGxBiao.cmdNew.ToolTipText, 1) '��������
mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
mod1.cmd.Parameters("@lcRen") = mod1.DName
mod1.cmd.Parameters("@lcUid") = mod1.DHid
mod1.cmd.Parameters("@nLb") = frmGxBiao.cmdNew.Tag
mod1.cmd.Execute
frmWBXJ.lblBid.Caption = mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblBh.Caption = "XJD" & mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblOid.Caption = mod1.cmd.Parameters("@bid").Value
frmWBXJ.lblLcou.Caption = Right(frmGxBiao.cmdNew.ToolTipText, 1) '��������
frmWBXJ.lblLc.Caption = 0
frmWBXJ.lblLcRen.Caption = mod1.DName
frmWBXJ.lblLcUid.Caption = mod1.DHid
frmWBXJ.lblNlb.Caption = frmGxBiao.cmdNew.Tag
frmWBXJ.lblYwy.Caption = mod1.DName
frmWBXJ.lblUid.Caption = mod1.DHid
frmWBXJ.lblBM.Caption = mod1.Bm
frmWBXJ.lblQy.Caption = mod1.Qy
frmWBXJ.lblZl.Caption = "ά��"
Set cmd = Nothing
If frmWBXJ.lblBh.Caption = "" Then
    ii = MsgBox("ϵͳ������������,�����̹ر�!�ٴδ򿪺�����Ϣ,������˴���.", vbOKOnly + vbExclamation, "A������")
    End
End If
'������Ŀ������Ϣ
tt = "select xmmc,xid from xmzl where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "'"
frmWBXJ.adoXm.Close
frmWBXJ.adoXm.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.comXmmc.RowSource = frmWBXJ.adoXm
frmWBXJ.comXmmc.ListField = "xmmc"
frmWBXJ.comXmmc.BoundColumn = "xid"

'tt = "select jzpb,pbid from bjxt_jzpb"
'frmWBXJ.adoPb.Close
'frmWBXJ.adoPb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Set frmWBXJ.comPb.RowSource = frmWBXJ.adoPb
'frmWBXJ.comPb.ListField = "jzpb"
'frmWBXJ.comPb.BoundColumn = "pbid"
            frmWBXJ.frmDx.Visible = False
            frmWBXJ.frmNb.Visible = True
            frmWBXJ.frmTime.Visible = True

            frmWBXJ.cmdD.Visible = True
            frmWBXJ.cmdJi.Visible = True
            frmWBXJ.tabGc.TabVisible(2) = False
            frmWBXJ.tabGc.TabVisible(0) = True
            frmWBXJ.tabGc.TabVisible(1) = True
            frmWBXJ.tabGc.Tab = 0

    '�������̰�ť
    Call modBJD.XJWBLcBut(44)
    
        frmWBXJ.cmdD.Visible = True

        frmWBXJ.cmdJi.Visible = True
    
frmWait.Visible = False
frmWBXJ.Visible = True
frmWBXJ.cmdMod.Enabled = False
'ˢ��ά�������б�
tt = "select * from xunJIaWbView where wbx='�걣' and bid=" & Val(frmWBXJ.lblBid.Caption)
    frmWBXJ.adoWb.Close
    frmWBXJ.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWBXJ.dtgWb.DataSource = frmWBXJ.adoWb
tt = "select * from xunJIaWbView where wbx='����' and bid=" & Val(frmWBXJ.lblBid.Caption)
    frmWBXJ.adoLj.Close
    frmWBXJ.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWBXJ.dtgLj.DataSource = frmWBXJ.adoLj
    frmWBXJ.cmdSave.Enabled = True
frmGxBiao.Enabled = False

'������Ϣ��
frmWBXJ.frmNew.Visible = True
tt = "select jzpb as ����Ʒ��,jzxh as �����ͺ�,sl as ����,jxId from wbjb where bid=" & Val(frmWBXJ.lblBid.Caption)
Set frmWBXJ.adoA = CreateObject("adodb.recordset")
frmWBXJ.adoA.Close
frmWBXJ.adoA.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmWBXJ.dtgA.DataSource = frmWBXJ.adoA

frmWBXJ.cmdBjd.Visible = False
frmWBXJ.txtHg.Locked = True
frmWBXJ.txtYhg.Locked = True
frmWBXJ.txtClf.Locked = True
frmWBXJ.cmdCG.Enabled = False
'frmWBXJ.cmdCong.Visible = False
frmWBXJ.cmdTk.Visible = True
End Sub

Private Sub cmdNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblZM.Caption = "    ָ��˾��ͻ�ǩ���ġ�����˾�Ĺ�����Ա��һ����ʱ�䷶Χ�ڣ�������9���£����豸����ά�������ĺ�ͬ�����Ժ���������ְ�����Ҳ����һ����ά�޸���ͻ�����6�����ʱ��ڵĺ�ͬ��"
End Sub


Private Sub cmdPj_Click()
Dim pk As String
Set frmLingPei.LpXh = CreateObject("adodb.recordset")
'Set frmLingPei.adoLpg = CreateObject("adodb.recordset")
Dim tt As String
Dim oo As Integer
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
    Exit Sub
End If

frmLingPei.Show


tt = "lpg_pei('')"
frmLingPei.adoLpg.Recordset.Close
frmLingPei.adoLpg.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
'Set frmLingPei.mga.DataSource = frmLingPei.adoLpg
Set frmLingPei.dtgView.DataSource = frmLingPei.adoLpg
 
''Set frmLingPei.mgc.DataSource = frmLingPei.adoLpg
'pk = "<        |<      ��  ��          |<  Ʒ  ��     |<  ��  ��           |< ��  ��     |< ��  ��  |< �����ۼ�    |<   �ɱ���   |<  ��  ��    "
'frmLingPei.mgb.FormatString = pk
If mod1.VLP = 1 Then
    frmLingPei.dtgView.Columns("�׼�").Visible = False
ElseIf mod1.VLP = 2 Then
    frmLingPei.dtgView.Columns("�׼�").Visible = False
ElseIf mod1.VLP = 3 Then
    frmLingPei.dtgView.Columns("�׼�").Visible = True
End If

If mod1.DName = "�Ŵ���" Then
    frmLingPei.cmdKq.Visible = True
    frmLingPei.frmMod.Visible = True
Else
    'frmLingPei.cmdKq.Visible = False
    'frmLingPei.frmMod.Visible = False
End If
'    Set frmlingpei.dtgView.DataSource = Nothing
End Sub

Private Sub cmdQH_Click()
If cmdQH.Caption = "��ʾ����" Then
    tt = "select * from xunjiaView1 where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "' order by bid desc"
    cmdQH.Caption = "��ʾ��Ч"
    cmdZF.Caption = "�ָ�"
Else
    tt = "select * from xunjiaView where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "' order by bid desc"
    cmdQH.Caption = "��ʾ����"
    cmdZF.Caption = "����"
End If

    frmGxBiao.adoXj.Close
    frmGxBiao.adoXj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
    If frmGxBiao.adoXj.RecordCount > 1 Then
        frmGxBiao.dtgXj.FixedRows = 0
        frmGxBiao.dtgXj.MergeCol(1) = True
        frmGxBiao.dtgXj.MergeCol(3) = True
        frmGxBiao.dtgXj.MergeCol(4) = True
        frmGxBiao.dtgXj.MergeCells = 3
        frmGxBiao.dtgXj.FixedRows = 1
    End If
End Sub

Private Sub cmdTl_Click()
Set frmLingjian.LpXh = CreateObject("adodb.recordset")
Dim tt As String
Dim oo As Integer
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
    Exit Sub
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

Private Sub cmdW_Click()
Call frmGxBNew.Bound
frmGxBNew.Show
frmGxBNew.ZOrder 0
End Sub

Private Sub cmdYk_Click()
Set frmLingjian.LpXh = CreateObject("adodb.recordset")
Dim oo As Integer
Dim tt As String
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
    Exit Sub
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

Private Sub cmdZF_Click()
'''''Dim tt As String
'''''Dim ZL As String
'''''Dim Bid As Long
'''''
'''''On Error Resume Next
'''''mod1.BTZ = 36
'''''If tabV.Tab = 0 Then
'''''    dtgXj.Col = 4
'''''    ZL = dtgXj.Text
'''''    dtgXj.Col = 6
'''''    Bid = dtgXj.Text
'''''Else
'''''    dtgGc.Col = 4
'''''    ZL = dtgGc.Text
'''''    dtgGc.Col = 6
'''''    Bid = dtgGc.Text
'''''
'''''End If
'''''
'''''If cmdZF.Caption = "����" Then
'''''    tt = "update xunjiad set delf=0 where bid=" & Bid
'''''Else
'''''    tt = "update xunjiad set delf=1 where bid=" & Bid
'''''End If
'''''Set mod1.HTP = CreateObject("adodb.recordset")
'''''mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'''''
''''''���Ϻ�,ԭ���������б�����ʧ.
'''''If cmdZF.Caption = "����" Then
'''''        tt = "update newfuwu set cf=1 where lx='ѯ�۵�' and bh='" & Bid & "'"
'''''    Set mod1.HTP = CreateObject("adodb.recordset")
'''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'''''End If
'''''
'''''frmGxBiao.adoXj.Requery
'''''Set frmGxBiao.dtgXj.DataSource = frmGxBiao.adoXj
'''''If frmGxBiao.adoXj.RecordCount > 1 Then
'''''    frmGxBiao.dtgXj.FixedRows = 0
'''''    frmGxBiao.dtgXj.MergeCol(1) = True
'''''    frmGxBiao.dtgXj.MergeCol(2) = True
'''''    frmGxBiao.dtgXj.MergeCol(3) = True
'''''    frmGxBiao.dtgXj.MergeCol(4) = True
'''''    frmGxBiao.dtgXj.MergeCol(5) = True
'''''    frmGxBiao.dtgXj.MergeCells = 3
'''''    frmGxBiao.dtgXj.FixedRows = 1
'''''End If
End Sub

Private Sub cmdZlG_Click()
MsgBox "ע:���ϼ۸���Ч����2005��11��20��"
Set frmLingjian.LpXh = CreateObject("adodb.recordset")
Dim oo As Integer
Dim tt As String
On Error Resume Next

frmZu.Enabled = False
If mod1.VLP = 0 Then
    Call mod1.NoQuan
    Exit Sub
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

Private Sub Command1_Click()

End Sub

Private Sub dtgXJ_DblClick()
Static Px As Boolean

If dtgXj.Row = 1 Then
    If Px = True Then
        dtgXj.Sort = 2
        Px = False
    Else
        dtgXj.Sort = 1
        Px = True
    End If
'Else
'    MsgBox MGa.ColData(1)
End If
End Sub


Private Sub dtgXj_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static ZF As Boolean
If Button <> 2 Then Exit Sub
If ZF = False Then
        dtgXj.FixedRows = 0
        dtgXj.MergeCol(1) = True
        dtgXj.MergeCol(3) = True
        dtgXj.MergeCol(4) = True
        dtgXj.MergeCells = 0
        dtgXj.FixedRows = 1
        ZF = True
Else
        dtgXj.FixedRows = 0
        dtgXj.MergeCol(1) = True
        dtgXj.MergeCol(3) = True
        dtgXj.MergeCol(4) = True
        dtgXj.MergeCells = 3
        dtgXj.FixedRows = 1
        ZF = False
End If
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
frmGxBiao.Width = mod1.FWidth
frmGxBiao.Height = mod1.FHeight
Set adoXj = CreateObject("adodb.recordset")
Set adoGc = CreateObject("adodb.recordset")
dtgXj.FixedCols = 0
dtgXj.Cols = 17
'''''''''If mod1.BM = "�����ҵ��" Then
'''''''''    dtgXj.ColWidth(0) = 6990
'''''''''    dtgXj.ColWidth(1) = 3500
'''''''''    dtgXj.ColWidth(2) = 2190
'''''''''    dtgXj.ColWidth(5) = 0
''''''''''''''    dtgXj.ColWidth(7) = 0
''''''''''''''    dtgXj.ColWidth(8) = 0
''''''''''''''    dtgXj.ColWidth(9) = 0
''''''''''''''    dtgXj.ColWidth(10) = 0
''''''''''''''    dtgXj.ColWidth(11) = 2500
''''''''''''''    dtgXj.ColWidth(12) = 1200
''''''''''''''    dtgXj.ColWidth(13) = 1950
'''''''''Else
'''''''''    dtgXj.ColWidth(0) = 6990
'''''''''    dtgXj.ColWidth(1) = 3500
'''''''''    dtgXj.ColWidth(6) = 0
'''''''''    dtgXj.ColWidth(7) = 0
'''''''''    dtgXj.ColWidth(8) = 0
'''''''''    dtgXj.ColWidth(9) = 0
'''''''''    dtgXj.ColWidth(10) = 0
'''''''''    dtgXj.ColWidth(11) = 0
'''''''''    dtgXj.ColWidth(12) = 0
'''''''''    dtgXj.ColWidth(13) = 0
'''''''''    dtgXj.ColWidth(14) = 0
'''''''''    dtgXj.ColWidth(15) = 0
'''''''''    dtgXj.ColWidth(16) = 0
'''''''''End If

    dtgXj.ColWidth(0) = 6990
    dtgXj.ColWidth(1) = 1000
    dtgXj.ColWidth(6) = 0
    dtgXj.ColWidth(7) = 0
    dtgXj.ColWidth(8) = 0
    dtgXj.ColWidth(9) = 0
    dtgXj.ColWidth(10) = 0
    dtgXj.ColWidth(11) = 0
    dtgXj.ColWidth(12) = 0
    dtgXj.ColWidth(13) = 0
    dtgXj.ColWidth(14) = 0
    dtgXj.ColWidth(15) = 0
    dtgXj.ColWidth(16) = 0
    dtgXj.ColWidth(17) = 0
If mod1.DName = "" Or mod1.DName = "��Ծ" Or Ywy = "�����" Then
    cmdW.Visible = True
Else
    cmdW.Visible = False
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
Cancel = True
frmGxBiao.Visible = False
frmZu.Enabled = True
End If
End Sub


Private Sub OKButton_Click()
Dim tt As String
Dim ZL As String
Dim Bid As Long
Dim htRow As Integer
dtgXj.Col = 3
ZL = dtgXj.Text

dtgXj.Cols = 17
dtgXj.Col = 16
htRow = Val(dtgXj.Text)
If htRow > 0 Or ZL = "ѯ��ָ��" Then
    dtgXj.Col = 5
    Bid = Val(dtgXj.Text)
    If Bid = 0 Then Exit Sub
    Call FmxcXJ.Bound(Bid)
    FmxcXJ.Show
    FmxcXJ.ZOrder 0
Exit Sub
End If

On Error Resume Next
mod1.BTZ = 36

If tabV.Tab = 0 Then
    dtgXj.Col = 3
    ZL = dtgXj.Text
    dtgXj.Col = 5
    Bid = dtgXj.Text
    'Exit Sub
Else
    dtgGc.Col = 4
    ZL = dtgGc.Text
    dtgGc.Col = 6
    Bid = dtgGc.Text

End If
    If Bid = 0 Then Exit Sub
Me.Enabled = False
frmWBXJ.Visible = False
frmGXBj.Visible = False
frmWait.Visible = True
frmWait.ZOrder 0
frmWait.Refresh
If ZL = "�˹�" Or ZL = "ά��" Or ZL = "����" Or ZL = "���̷ְ�" Then
    dtgXj.Col = 5
    Bid = Val(dtgXj.Text)
            Call frmWBXX.Qing
            Call frmWBXX.Bound(Bid)
            'Call frmWBXNew.Bound(Val(dtgFL.Text))
            frmWBXX.Show
            frmWBXX.ZOrder 0
    
        frmGxBiao.Enabled = True
Exit Sub
End If

If ZL = "ά��" Or ZL = "����" Or ZL = "���̷ְ�" Then
    If mod1.Bm = "�����ҵ��" Or mod1.Qy = "����" Then
            Call modBJD.BJDGXQing
            Call modBJD.BJDGDBound(Bid)
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
    Else
        Call modBJD.BJDWBQing
        Call modBJD.BJDGXQing
        Call modBJD.BJDBound(Bid, ZL)
        Call modBJD.wbxjLocked
        tt = "select bid from xunjiaOld where oid=" & Val(frmWBXJ.lblOid.Caption) & " order by bid"
        frmWBXJ.adoOid.Close
        frmWBXJ.adoOid.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        If frmWBXJ.adoOid.RecordCount > 1 Then
            frmWBXJ.cmdRight.Enabled = False
            frmWBXJ.cmdLeft.Enabled = True
        Else
            frmWBXJ.cmdRight.Enabled = False
            frmWBXJ.cmdRight.Enabled = False
        End If
   
        frmWait.Visible = False
        frmWBXJ.Visible = True
        frmWBXJ.ZOrder 0
        frmWBXJ.cmdMod.Enabled = True
        frmWBXJ.cmdSave.Enabled = False
        frmWBXJ.adoOid.MoveLast
    End If
ElseIf (ZL = "����" Or ZL = "���" Or ZL = "��Ʒ" Or ZL = "�����" Or ZL = "���ѯ�۵�") Then
    Call modBJD.BJDWBQing
    Call modBJD.BJDGXQing
    Call modBJD.BJDBound(Bid, ZL)
    Call modBJD.gxbjLocked
'''''''''    tt = "select bid from xunjiaOld where oid=" & Val(frmGXBj.lblOid.Caption) & " order by bid"
'''''''''    frmGXBj.adoOid.Close
'''''''''    frmGXBj.adoOid.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''    frmGXBj.adoOid.MoveLast
'''''''''    If frmGXBj.adoOid.RecordCount > 1 Then
'''''''''        frmGXBj.cmdRight.Enabled = False
'''''''''        frmGXBj.cmdLeft.Enabled = True
'''''''''    Else
'''''''''        frmGXBj.cmdRight.Enabled = False
'''''''''        frmGXBj.cmdLeft.Enabled = False
'''''''''    End If

            Call frmGXBj.dtgMaFF
            Call modBJD.gxbjLocked
            If frmGXBj.lblYwy = "лѩ÷" Or Bid > 10058 Then
                'frmGXBj.frmSD.Visible = True
                frmGXBj.frmCg.Top = 4740
                frmGXBj.dtgNew.Visible = True
                
                frmGXBj.dtgP.Visible = True
            Else
                'frmGXBj.frmSD.Visible = False
                frmGXBj.frmCg.Top = 7620
                frmGXBj.dtgNew.Visible = False

                frmGXBj.dtgP.Visible = False
            End If
    frmWait.Visible = False
    frmGXBj.Visible = True
    frmGXBj.ZOrder 0
    frmGXBj.cmdMod.Enabled = True
    frmGXBj.cmdSave.Enabled = False
End If
End Sub



Public Sub XJBound(Ra, La As Integer)
Dim oo As Integer: Dim ii As Integer
Dim tt As String
dtgXj.Visible = False
dtgXj.Clear
dtgXj.Rows = La + 1
dtgXj.Row = 0
dtgXj.Col = 0: dtgXj.Text = "��Ŀ����": dtgXj.CellFontBold = True
dtgXj.Col = 1: dtgXj.Text = "�˹���": dtgXj.CellFontBold = True
dtgXj.Col = 2: dtgXj.Text = "ѯ������": dtgXj.CellFontBold = True
dtgXj.Col = 3: dtgXj.Text = "ҵ��Ա": dtgXj.CellFontBold = True
dtgXj.Col = 4: dtgXj.Text = "���": dtgXj.CellFontBold = True
For oo = 1 To La
    dtgXj.Row = oo
    For ii = 0 To 5
        dtgXj.Col = ii
        If IsNull(Ra(ii, oo - 1)) = True Then
            dtgXj.Text = ""
        Else
            dtgXj.Text = Ra(ii, oo - 1)
        End If
    Next
Next
        'tt = "select xmmc as ��Ŀ����,yhg as �˹���,rq as ѯ������,ywy as ҵ��Ա,BianHao AS ���,bid from xunJiaD where zl='�˹�' and lc>=2 order by rq desc"
dtgXj.Visible = True
End Sub
