VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGZDBR 
   Caption         =   "��������ѯ"
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
         Caption         =   "�����ල��"
         Height          =   285
         Left            =   1050
         TabIndex        =   22
         Top             =   150
         Width           =   1245
      End
      Begin VB.CommandButton cmdFw 
         Caption         =   "ѡ��ҵ��Ա"
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
         Caption         =   "��ѯ"
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
         Text            =   "���"
         Top             =   0
         Width           =   945
      End
      Begin VB.CommandButton cmdAll2 
         Caption         =   "ȫ  ��"
         Height          =   285
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label5 
         Caption         =   "ֵ"
         Height          =   315
         Left            =   1830
         TabIndex        =   21
         Top             =   30
         Width           =   315
      End
      Begin VB.Label Label6 
         Caption         =   "��ѯ��ʽ"
         Height          =   285
         Left            =   60
         TabIndex        =   20
         Top             =   0
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "ȫ  ��"
      Height          =   285
      Left            =   90
      TabIndex        =   11
      Top             =   8880
      Width           =   9405
   End
   Begin VB.CommandButton cmdOpen2 
      Caption         =   "��    ��"
      Height          =   255
      Left            =   12240
      TabIndex        =   10
      Top             =   30
      Width           =   2955
   End
   Begin VB.CommandButton cmdOpen1 
      Caption         =   "��    ��"
      Height          =   285
      Left            =   3720
      TabIndex        =   9
      Top             =   30
      Width           =   5775
   End
   Begin VB.CommandButton cmdREF 
      Caption         =   "��ѯ"
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
      Text            =   "���"
      Top             =   8520
      Width           =   2715
   End
   Begin VB.CommandButton cmdback 
      Caption         =   "����"
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
      Caption         =   "ֵ"
      Height          =   315
      Left            =   4110
      TabIndex        =   7
      Top             =   8550
      Width           =   315
   End
   Begin VB.Label Label3 
      Caption         =   "��ѯ��ʽ"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   8520
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "δ��ɹ�����"
      ForeColor       =   &H00FF00FF&
      Height          =   195
      Left            =   10260
      TabIndex        =   1
      Top             =   60
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "��������¼"
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
Public adoY As ADODB.Recordset '����ɵĹ�����
Public adoW As ADODB.Recordset 'δ��ɵĹ�����

Private Sub cmdAll_Click()
On Error Resume Next
Dim tt As String
tt = "select ��������,���,����������,gid,fl FROM gzdView where trq is null and ҵ��Ա='" & lblFw.Caption & "' and uid='" & lblFw.ToolTipText & "' order by ��������"
frmGZDBR.adoW.Close
frmGZDBR.adoW.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgW.DataSource = frmGZDBR.adoW

tt = "Select ��������,���,����������,ҵ��Ա,gid,uid,qy,trq,��Ŀ����,����,fl,�ϸ�� from gzdView where not(trq is null) and ҵ��Ա='" & lblFw.Caption & "' and uid='" & lblFw.ToolTipText & "' order by gid desc"
frmGZDBR.adoY.Close
frmGZDBR.adoY.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgY.DataSource = frmGZDBR.adoY
frmGZDBR.dtgY.Row = 1

'Set frmGZDBR.dtgY.DataSource = frmGZDBR.adoW
End Sub

Private Sub cmdAll2_Click()
On Error Resume Next
Dim tt As String
tt = "select ��������,���,����������,gid,fl FROM gzdView where trq is null and ҵ��Ա='" & mod1.DName & "' and uid='" & mod1.DHid & "' order by ��������"
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
If comLx.Text = "���" Then
    If mod1.DName = "����" Or mod1.DName = "����" Or mod1.DName = "����" Or mod1.DName = "������" Or mod1.DName = "Ǯب" Then
        tt = "Select * from gzdView where not(trq is null) and ���=" & Val(txtZ.Text) & "  order by gid desc"
    ElseIf mod1.KhK = 1 Then
        tt = "Select * from gzdView where not(trq is null) and ���=" & Val(txtZ.Text) & " and bm='" & mod1.BM & "' order by gid desc"
    Else
        tt = "Select * from gzdView where not(trq is null) and ���=" & Val(txtZ.Text) & " and ҵ��Ա='" & mod1.DName & "' order by gid desc"
    End If

ElseIf comLx.Text = "��Ŀ����" Then
'    If mod1.DName = "Ǯب" Or mod1.DName = "����" Or mod1.DName = "������" Or mod1.DName = "����" Or mod1.DName = "������" Then
'        tt = "Select * from gzdView where not(trq is null) and ��Ŀ���� like '%" & txtZ.Text & "%' and ҵ��Ա='" & mod1.DName & "' order by gid desc"
'    ElseIf mod1.KhK = 1 Then
'        tt = "Select * from gzdView where not(trq is null) and ��Ŀ���� like '%" & txtZ.Text & "%' and bm='" & mod1.DName & "' order by gid desc"
'    ElseIf mod1.DName = "Ǯب" Or mod1.DName = "����" Or mod1.DName = "������" Or mod1.DName = "����" Or mod1.DName = "������" Then
'        tt = "Select * from gzdView where not(trq is null) and ��Ŀ���� like '%" & txtZ.Text & "%'  order by gid desc"
'    End If

    If mod1.DName = "����" Or mod1.DName = "����" Or mod1.DName = "����" Or mod1.DName = "������" Or mod1.DName = "Ǯب" Then
        tt = "Select * from gzdView where not(trq is null) and ��Ŀ���� like '%" & txtZ.Text & "%'  order by gid desc"
    ElseIf mod1.KhK = 1 Then
        tt = "Select * from gzdView where not(trq is null) and ��Ŀ���� like '%" & txtZ.Text & "%' and bm='" & mod1.BM & "' order by gid desc"
    Else
        tt = "Select * from gzdView where not(trq is null) and ��Ŀ���� like '%" & txtZ.Text & "%' and ҵ��Ա='" & mod1.DName & "' order by gid desc"
    End If
ElseIf comLx.Text = "���" Then
    If mod1.DName = "����" Or mod1.DName = "����" Or mod1.DName = "����" Or mod1.DName = "������" Or mod1.DName = "Ǯب" Then
        tt = "Select * from gzdView where not(trq is null) and ���������� like '%" & txtZ.Text & "%'  order by gid desc"
    ElseIf mod1.KhK = 1 Then
        tt = "Select * from gzdView where not(trq is null) and ���������� like '%" & txtZ.Text & "%' and bm='" & mod1.BM & "' order by gid desc"
    Else
        tt = "Select * from gzdView where not(trq is null) and ���������� like '%" & txtZ.Text & "%' and ҵ��Ա='" & mod1.DName & "' order by gid desc"
    End If
End If
frmGZDBR.adoY.Close
frmGZDBR.adoY.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgY.DataSource = frmGZDBR.adoY

If comLx.Text = "���" Then
    tt = "Select ��������,���,����������,gid,fl from gzdView where trq is null and ���=" & Val(txtZ.Text) & " order by gid desc"
ElseIf comLx.Text = "��Ŀ����" Then
    tt = ""
   ' tt = "Select * from gzdView where trq is null and ��Ŀ���� like '%" & txtZ.Text & "%' and ҵ��Ա='" & mod1.DName & "' order by gid desc"
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

tt = "Select  ��������,���,����������,gid,fl  from gzdView where trq is null and ���=" & Val(txtW.Text) & " order by gid desc"

frmGZDBR.adoW.Close
frmGZDBR.adoW.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgW.DataSource = frmGZDBR.adoW
End Sub


Private Sub cmdZJ_Click()
Dim tt As String
On Error Resume Next
lblFw.Caption = "Ǯب"
lblFw.ToolTipText = "HM152"

tt = "Select * from gzdView where not(trq is null) and ҵ��Ա='" & lblFw.Caption & "' and uid='" & lblFw.ToolTipText & "' order by gid desc"
frmGZDBR.adoY.Close
frmGZDBR.adoY.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
Set frmGZDBR.dtgY.DataSource = frmGZDBR.adoY

tt = "select ��������,���,����������,gid,fl FROM gzdView where trq is null and ҵ��Ա='" & lblFw.Caption & "' and uid='" & lblFw.ToolTipText & "' order by ��������"
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
