VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form fyBB 
   Caption         =   "���ñ���"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Frame frmGc 
      Caption         =   "���̲���Աѡ��"
      Height          =   6705
      Left            =   30
      TabIndex        =   48
      Top             =   30
      Width           =   4785
      Begin VB.CommandButton cmdXz 
         Caption         =   "�ر�"
         Height          =   285
         Left            =   3690
         TabIndex        =   52
         Top             =   6390
         Width           =   825
      End
      Begin MSDataListLib.DataList dtGC 
         Height          =   5940
         Left            =   2880
         TabIndex        =   51
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   10478
         _Version        =   393216
      End
      Begin VB.CommandButton cmdQuan 
         Caption         =   "ȫ��"
         Height          =   285
         Left            =   150
         TabIndex        =   50
         Top             =   6390
         Width           =   885
      End
      Begin VB.CommandButton cmdZu 
         Caption         =   "���"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   49
         Top             =   390
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdXuan 
      Caption         =   "ѡ ȡ"
      Height          =   285
      Left            =   14250
      TabIndex        =   47
      Top             =   7350
      Width           =   945
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "�� ��"
      Height          =   285
      Left            =   14250
      TabIndex        =   46
      Top             =   7650
      Width           =   945
   End
   Begin VB.Frame frmLb 
      Caption         =   "��ѯ��Ŀ"
      Height          =   2295
      Left            =   4110
      TabIndex        =   8
      Top             =   6900
      Width           =   9975
      Begin VB.CommandButton cmdYwy 
         Caption         =   "ҵ��Ա"
         Height          =   315
         Left            =   2790
         TabIndex        =   44
         Top             =   1920
         Width           =   1005
      End
      Begin VB.CommandButton cmdQing 
         Caption         =   "ȫ��"
         Height          =   315
         Left            =   1470
         TabIndex        =   43
         Top             =   1920
         Width           =   945
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "ȫѡ"
         Height          =   315
         Left            =   150
         TabIndex        =   42
         Top             =   1920
         Width           =   1005
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "����"
         Height          =   255
         Index           =   32
         Left            =   5400
         TabIndex        =   41
         Top             =   1560
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "�׺�"
         Height          =   255
         Index           =   31
         Left            =   4095
         TabIndex        =   40
         Top             =   1560
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "���߷�"
         Height          =   255
         Index           =   30
         Left            =   2790
         TabIndex        =   39
         Top             =   1560
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "����������"
         Height          =   255
         Index           =   29
         Left            =   1485
         TabIndex        =   38
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "����ͣ����"
         Height          =   255
         Index           =   28
         Left            =   180
         TabIndex        =   37
         Top             =   1560
         Width           =   1245
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "�Ŷӽ����"
         Height          =   255
         Index           =   27
         Left            =   8010
         TabIndex        =   36
         Top             =   1260
         Width           =   1215
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "����������"
         Height          =   255
         Index           =   26
         Left            =   6705
         TabIndex        =   35
         Top             =   1260
         Width           =   1215
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "��ѵ��"
         Height          =   255
         Index           =   25
         Left            =   5400
         TabIndex        =   34
         Top             =   1260
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "��ݷ�"
         Height          =   255
         Index           =   24
         Left            =   4095
         TabIndex        =   33
         Top             =   1260
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "��Ա��Ƹ"
         Height          =   255
         Index           =   23
         Left            =   2790
         TabIndex        =   32
         Top             =   1260
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "�г��ƹ�"
         Height          =   255
         Index           =   22
         Left            =   1485
         TabIndex        =   31
         Top             =   1260
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "����"
         Height          =   255
         Index           =   21
         Left            =   180
         TabIndex        =   30
         Top             =   1260
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "�칫��Ʒ"
         Height          =   255
         Index           =   20
         Left            =   8010
         TabIndex        =   29
         Top             =   930
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "�绰"
         Height          =   255
         Index           =   19
         Left            =   6705
         TabIndex        =   28
         Top             =   930
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "ˮ��"
         Height          =   255
         Index           =   18
         Left            =   5400
         TabIndex        =   27
         Top             =   930
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "��ҵ��"
         Height          =   255
         Index           =   17
         Left            =   4095
         TabIndex        =   26
         Top             =   930
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "����"
         Height          =   255
         Index           =   16
         Left            =   2790
         TabIndex        =   25
         Top             =   930
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "��Ʒ��"
         Height          =   255
         Index           =   15
         Left            =   1485
         TabIndex        =   24
         Top             =   930
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "�д���"
         Height          =   255
         Index           =   14
         Left            =   180
         TabIndex        =   23
         Top             =   930
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "�ͷ�"
         Height          =   255
         Index           =   13
         Left            =   8010
         TabIndex        =   22
         Top             =   600
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "ס�޷�"
         Height          =   255
         Index           =   12
         Left            =   6705
         TabIndex        =   21
         Top             =   600
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "�����Ŷӷ�"
         Height          =   255
         Index           =   11
         Left            =   5400
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "�˷�"
         Height          =   255
         Index           =   10
         Left            =   4095
         TabIndex        =   19
         Top             =   600
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "������"
         Height          =   255
         Index           =   9
         Left            =   2790
         TabIndex        =   18
         Top             =   600
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "ͣ����"
         Height          =   255
         Index           =   8
         Left            =   1485
         TabIndex        =   17
         Top             =   600
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "���⽻ͨ��"
         Height          =   255
         Index           =   7
         Left            =   180
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "���ڽ�ͨ��"
         Height          =   255
         Index           =   6
         Left            =   8010
         TabIndex        =   15
         Top             =   270
         Width           =   1275
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "ͨ�ŷ�"
         Height          =   255
         Index           =   5
         Left            =   6705
         TabIndex        =   14
         Top             =   270
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "���·�"
         Height          =   255
         Index           =   4
         Left            =   5400
         TabIndex        =   13
         Top             =   270
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "���η�"
         Height          =   255
         Index           =   3
         Left            =   4095
         TabIndex        =   12
         Top             =   270
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "���ݲ���"
         Height          =   255
         Index           =   2
         Left            =   2790
         TabIndex        =   11
         Top             =   270
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "�Ľ�"
         Height          =   255
         Index           =   1
         Left            =   1485
         TabIndex        =   10
         Top             =   270
         Width           =   1035
      End
      Begin VB.CheckBox chkLb 
         Caption         =   "��ͬ���"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9990
         Y1              =   1860
         Y2              =   1860
      End
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14790
      Picture         =   "fyBB.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "����"
      Top             =   8790
      Width           =   465
   End
   Begin VB.CommandButton cmdCX 
      Caption         =   "�� ѯ"
      Height          =   315
      Left            =   14250
      TabIndex        =   3
      Top             =   7020
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgFybb 
      Height          =   6795
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "˫���ɴ򿪴˱�����"
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   11986
      _Version        =   393216
      BackColorBkg    =   8421504
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdFw 
      Caption         =   "ѡ��Ա������"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   7050
      Width           =   1425
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   315
      Left            =   1500
      TabIndex        =   5
      Top             =   7560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   8454016
      Format          =   81657857
      CurrentDate     =   38797
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   315
      Left            =   1500
      TabIndex        =   6
      Top             =   7950
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   8454016
      Format          =   81657857
      CurrentDate     =   38797
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   735
      Left            =   210
      TabIndex        =   45
      Top             =   8310
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Label Label1 
      Caption         =   "����:"
      Height          =   225
      Left            =   810
      TabIndex        =   4
      Top             =   7590
      Width           =   465
   End
   Begin VB.Label lblFw 
      Height          =   225
      Left            =   1530
      TabIndex        =   1
      Top             =   7110
      Width           =   2475
   End
End
Attribute VB_Name = "fyBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoFyBB As Object
Dim Lb As String 'ѡ���ѯ����Ŀ
Dim LX As String '��Ӧ��ѯ��Ŀ�ļ�¼�������0
Dim Fw As String '��ѯ��Ա��Χ
Dim adoGc As Object

Private Sub cmdAll_Click()
Dim oo As Integer
For oo = 0 To 32
    chkLb(oo).Value = 1
Next
End Sub

Private Sub cmdBack_Click()
fyBB.Visible = False
If frmBxV.Visible = True Then
    frmBxV.Enabled = True
    frmBxV.ZOrder 0
Else
    frmZu.Enabled = True
    frmZu.ZOrder 0
End If
End Sub

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText dtgFybb.Clip
dtgFybb.FixedRows = 1
End Sub

Private Sub cmdCx_Click()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
Dim FHg As Single
On Error Resume Next
If lblFw.Caption = "" Then
    MsgBox "��ѡ����Ӧ����Ա!"
    Exit Sub
End If
Lb = ""
LX = ""
frmGc.Visible = False
Me.Enabled = False
frmWait.Show
frmWait.ZOrder 0
Me.MousePointer = 11

    For oo = 0 To 32
        If chkLb(oo).Value = 1 Then
            Lb = Lb & "," & chkLb(oo).Caption
                If chkLb(oo).Caption <> "��ͬ���" Then
                If LX = "" Then
                    LX = chkLb(oo).Caption & ">0"
                Else
                    LX = LX & " or " & chkLb(oo).Caption & ">0"
                End If
            End If
        End If
    Next
    If LX <> "" Then
        LX = "and (" & LX & ")"
    End If
If lblFw.ToolTipText <> "" And Left((lblFw.ToolTipText), 1) = "H" Then 'ѡ����Ա

    tt = "select bxid,qy,bm,comid,����,����" & Lb & " from ����ͳ��A where ����='" & lblFw.Caption & "' and ywyuid='" & _
        lblFw.ToolTipText & "' and ����>='" & dt1.Value & "' and ����<'" & DateSerial(Year(dt2.Value), Month(dt2.Value), Day(dt2.Value) + 1) & _
        "'" & LX & " order by ����"
ElseIf lblFw.ToolTipText <> "" And Left((lblFw.ToolTipText), 1) <> "H" And Val(lblFw.ToolTipText) < 10 Then
    tt = "select bxid,qy,bm,comid,����,����" & Lb & " from ����ͳ��A where left(cast(gzu as nvarchar(3)),1)='" & _
    Left(lblFw.ToolTipText, 1) & "' and gzu<100 and ����>='" & dt1.Value & "' and ����<'" & DateSerial(Year(dt2.Value), Month(dt2.Value), Day(dt2.Value) + 1) & _
        "'" & LX & " order by ����"
ElseIf lblFw.ToolTipText <> "" And Left((lblFw.ToolTipText), 1) <> "H" And Val(lblFw.ToolTipText) > 10 Then
    tt = "select bxid,qy,bm,comid,����,����" & Lb & " from ����ͳ��A where gzu=" & _
    Val(lblFw.ToolTipText) & " and gzu<100 and ����>='" & dt1.Value & "' and ����<'" & DateSerial(Year(dt2.Value), Month(dt2.Value), Day(dt2.Value) + 1) & _
        "'" & LX & " order by ����"
ElseIf lblFw.Caption = "��ع��̲�" Then
    tt = "select bxid,qy,bm,comid,����,����" & Lb & " from ����ͳ��A where (left(cast(gzu as nvarchar(3)),1)=5 or left(cast(gzu as nvarchar(3)),1)=6 or gzu=4) and gzu<100 and ����>='" & dt1.Value & "' and ����<'" & DateSerial(Year(dt2.Value), Month(dt2.Value), Day(dt2.Value) + 1) & _
        "'" & LX & " order by ����"
ElseIf lblFw.Caption = "���̲�" Then
    tt = "select bxid,qy,bm,comid,����,����" & Lb & " from ����ͳ��A where bm='���̲�' and ����>='" & dt1.Value & "' and ����<'" & DateSerial(Year(dt2.Value), Month(dt2.Value), Day(dt2.Value) + 1) & _
        "'" & LX & " order by ����"
ElseIf lblFw.Caption = "��������" Then
    tt = "select bxid,qy,bm,comid,����,����" & Lb & " from ����ͳ��A where (bm='��������' or bm='����') and ����>='" & dt1.Value & "' and ����<'" & DateSerial(Year(dt2.Value), Month(dt2.Value), Day(dt2.Value) + 1) & _
        "'" & LX & " order by bm,����"
Else 'ѡ����

End If



adoFyBB.Close
adoFyBB.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgFybb.DataSource = adoFyBB
frmWait.Visible = True
frmWait.ZOrder 0
frmWait.Refresh
dtgFybb.Rows = adoFyBB.RecordCount + 2
dtgFybb.Row = dtgFybb.Rows - 1
dtgFybb.Col = 6
dtgFybb.Text = "�ϼ�"
For oo = 7 To dtgFybb.Cols - 1
    dtgFybb.Col = oo
    FHg = 0
    For ii = 1 To adoFyBB.RecordCount
        dtgFybb.Row = ii
        FHg = dtgFybb.Text + FHg
    Next
    dtgFybb.Row = dtgFybb.Rows - 1
    dtgFybb.Text = FHg
Next
dtgFybb.Rows = dtgFybb.Rows + 1
dtgFybb.Row = dtgFybb.Rows - 1
dtgFybb.Col = 6
dtgFybb.Text = "�ܼ�"
'�����ܼ�
FHg = 0
dtgFybb.Row = dtgFybb.Row - 1
For oo = 7 To dtgFybb.Cols - 1
    dtgFybb.Col = oo
    FHg = FHg + dtgFybb.Text


Next
dtgFybb.Row = dtgFybb.Row + 1
dtgFybb.Col = 7
dtgFybb.Text = FHg
Me.Enabled = True
frmWait.Visible = False
Me.ZOrder 0
    Me.MousePointer = 0
End Sub

Private Sub cmdFw_Click()
Dim ii As Integer
If mod1.Bm <> "���̲�" Then
    If mod1.KhK > 1 And mod1.DName <> "����" Then
        ii = MsgBox("�Ƿ��ѯ���̲���", vbYesNo + vbInformation + vbDefaultButton2)
        If ii = vbNo Then
            Set Ren.XForm = New fyBB
            Call mod1.RenXz("fyBB", Me, 0)
        Else
            frmGc.Visible = True
        End If
    Else
            Set Ren.XForm = New fyBB
            Call mod1.RenXz("fyBB", Me, 0)
    End If
Else
    frmGc.Visible = True
End If
End Sub

Private Sub cmdQing_Click()
Dim oo As Integer
For oo = 0 To 32
    chkLb(oo).Value = 0
Next
End Sub

Private Sub cmdQuan_Click()
lblFw.Caption = ""
lblFw.ToolTipText = ""

    If mod1.Zuf = 1 Then
        lblFw.Caption = mod1.DName & "��"
        lblFw.ToolTipText = mod1.Gzu
'    ElseIf mod1.DName = "֣��" Then
'        lblFw.Caption = "��ع��̲�"
    ElseIf mod1.DName = "����" Or mod1.Bm = "�ܾ���" Or mod1.Bm = "����" Or mod1.Bm = "ά����" Then
        lblFw.Caption = "���̲�"
    End If
End Sub

Private Sub cmdXuan_Click()
dtgFybb.FixedRows = 0
End Sub

Private Sub cmdXZ_Click()
frmGc.Visible = False
End Sub

Private Sub cmdZu_Click(Index As Integer)
Dim tt As String
On Error Resume Next
tt = "select username,userid from worker where gzu=" & cmdZu(Index).Tag & " order by zzf desc,zuf desc"
adoGc.Close
adoGc.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgC.RowSource = adoGc
dtgC.ListField = "username"
dtgC.BoundColumn = "userid"
lblFw.Caption = ""
lblFw.ToolTipText = ""
lblFw.Caption = cmdZu(Index).Caption & "��"
lblFw.ToolTipText = cmdZu(Index).Tag
End Sub

Private Sub dtgC_Click()
lblFw.Caption = ""
lblFw.ToolTipText = ""
lblFw.Caption = dtgC.Text
lblFw.ToolTipText = dtgC.BoundText
End Sub


Private Sub dtgFybb_Click()
frmGc.Visible = False
End Sub

Private Sub dtgFybb_DblClick()
dtgFybb.Col = 1
'MsgBox MGa.Text

If Val(dtgFybb.Text) = 0 Then Exit Sub
If mod1.DKZ(dtgFybb.Text, 2) = True Then
        MsgBox "��ݱ�����" & mod1.DKRen & "��,���Ժ�����,������������ϵ."
        Exit Sub
End If

Me.Enabled = False
frmFYBX.Show
mod1.BTZ = 23
Call ModBx.FyQing
Call ModBx.fydBound(Val(dtgFybb.Text))
frmFYBX.cmdSave.Enabled = False
frmFYBX.cmdMod.Enabled = False
End Sub


Private Sub Form_Click()
frmGc.Visible = False
End Sub

Private Sub Form_Load()
Dim oo As Integer
Dim tt As String
Dim zz As String '���鳤����
On Error Resume Next
dt2.Value = Date
Set adoFyBB = CreateObject("adodb.recordset")
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
dtgFybb.ColWidth(0) = 300
dtgFybb.ColWidth(1) = 0
dtgFybb.ColWidth(2) = 0
dtgFybb.ColWidth(3) = 0
dtgFybb.ColWidth(4) = 0
dtgFybb.ColWidth(6) = 2500
dtgFybb.ColWidth(40) = 0
For oo = 0 To 32
    chkLb(oo).Value = 1
Next
dt1.Value = DateSerial(Year(Date), 1, 1)
If mod1.Bm = "���̲�" Or mod1.Bm = "�ܾ���" Or mod1.Bm = "����" Or mod1.Bm = "ά����" Then
    For oo = 20 To 1 Step -1
        Unload cmdZu(oo)
    Next

    If mod1.Zuf = 1 Then
        tt = "select username,userid,gzu from worker where left(cast(gzu as nvarchar(3)),1)=" & Str(mod1.Gzu) & " and gzu<100 and zuf=1 order by zzf desc,gzu"
'    ElseIf mod1.DName = "֣��" Then
'        tt = "select username,userid,gzu from worker where (left(cast(gzu as nvarchar(3)),1)=5 or left(cast(gzu as nvarchar(3)),1)=6 or gzu=4) and (zuf=1 or zuf=2) order by zzf desc,left(cast(gzu as nvarchar(3)),1),gzu,zuf desc"
    ElseIf mod1.DName = "����" Or mod1.Bm = "�ܾ���" Or mod1.Bm = "����" Or mod1.Bm = "ά����" Then
        tt = "select username,userid,gzu from worker where zuf=1 and gzu<100 order by zzf desc,left(cast(gzu as nvarchar(3)),1),gzu,zuf desc"
    End If
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        mod1.HTP.MoveFirst
        oo = 0
        Do While Not mod1.HTP.EOF
            If oo > 0 Then
                Load cmdZu(oo)
                cmdZu(oo).Top = cmdZu(oo - 1).Top + 350
                cmdZu(oo).Visible = True
            End If
            If mod1.HTP.Fields("gzu").Value < 10 Then
                cmdZu(oo).Caption = mod1.HTP.Fields("username").Value
                zz = mod1.HTP.Fields("username").Value
            Else
                cmdZu(oo).Caption = zz & "(" & mod1.HTP.Fields("username").Value & ")"
            End If
            cmdZu(oo).ToolTipText = mod1.HTP.Fields("userid").Value
            cmdZu(oo).Tag = mod1.HTP.Fields("gzu").Value
            oo = oo + 1
            mod1.HTP.MoveNext
        Loop
End If
frmGc.Visible = False
Set adoGc = CreateObject("adodb.recordset")
End Sub

Private Sub Form_Resize()
frmGc.Visible = False
End Sub


