VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form frmKhbrG 
   Caption         =   "��Ŀ���ϲ�ѯ"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15180
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   15180
   Begin NiceFormControl.NiceButton NiceButton1 
      Height          =   285
      Left            =   13200
      TabIndex        =   30
      Top             =   7080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      BTYPE           =   1
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKhbrG.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Caption         =   "����˿ͻ�"
   End
   Begin NiceFormControl.NiceButton cmdLZ 
      Height          =   375
      Left            =   10950
      TabIndex        =   29
      Top             =   2220
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   661
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKhbrG.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Caption         =   "��ְ��Ա��Ŀ"
   End
   Begin VB.CommandButton cmdYwy 
      Caption         =   "ѡ�񻮹���"
      Height          =   315
      Left            =   13230
      TabIndex        =   27
      Top             =   7650
      Width           =   1185
   End
   Begin VB.CommandButton cmdHG 
      Caption         =   "��Ŀ����"
      Height          =   345
      Left            =   10470
      TabIndex        =   26
      Top             =   7650
      Width           =   1485
   End
   Begin VB.CommandButton cmdXq 
      Caption         =   "��  ��"
      Height          =   405
      Left            =   10950
      TabIndex        =   25
      Top             =   930
      Width           =   3405
   End
   Begin VB.CommandButton cmdFw 
      Caption         =   "��ѯ��Χ"
      Height          =   315
      Left            =   10800
      TabIndex        =   23
      Top             =   5100
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   585
      Left            =   14370
      Picture         =   "frmKhbrG.frx":0038
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8520
      Width           =   675
   End
   Begin MSDataListLib.DataCombo comYwy 
      Height          =   330
      Left            =   11850
      TabIndex        =   21
      Top             =   3540
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "ȫ��"
      Height          =   315
      Left            =   13590
      TabIndex        =   20
      Top             =   3510
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3105
      Left            =   5760
      TabIndex        =   12
      Top             =   300
      Visible         =   0   'False
      Width           =   4755
      Begin VB.CommandButton cmdNew 
         Caption         =   "�� �� �� Ŀ"
         Height          =   315
         Left            =   390
         TabIndex        =   16
         Top             =   1530
         Width           =   3975
      End
      Begin VB.ComboBox comHyXz 
         Height          =   300
         ItemData        =   "frmKhbrG.frx":013A
         Left            =   1860
         List            =   "frmKhbrG.frx":0153
         TabIndex        =   15
         Top             =   1020
         Width           =   2505
      End
      Begin VB.TextBox txtKhdm 
         Height          =   345
         Left            =   2460
         TabIndex        =   14
         Text            =   "khdm"
         Top             =   2310
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox comKhmc 
         Height          =   300
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   540
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.Label Label5 
         Caption         =   "�ͻ���ҵ���ʣ�"
         Height          =   285
         Left            =   420
         TabIndex        =   18
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "�ͻ����ƣ�"
         Height          =   315
         Left            =   780
         TabIndex        =   17
         Top             =   570
         Visible         =   0   'False
         Width           =   945
      End
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "frmKhbrG.frx":0195
      Left            =   11880
      List            =   "frmKhbrG.frx":019C
      TabIndex        =   9
      Top             =   5610
      Width           =   2595
   End
   Begin VB.TextBox txtZ 
      Height          =   315
      Left            =   11880
      TabIndex        =   8
      Top             =   6300
      Width           =   2595
   End
   Begin VB.CommandButton cmdVall 
      Caption         =   "��ʾȫ��"
      Height          =   375
      Left            =   10470
      TabIndex        =   7
      Top             =   7140
      Width           =   1485
   End
   Begin TabDlg.SSTab tabCx 
      Height          =   9165
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   16166
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "���ͻ���ѯ"
      TabPicture(0)   =   "frmKhbrG.frx":01AA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dtgKh"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmPx"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "����ϵ�˲�ѯ"
      TabPicture(1)   =   "frmKhbrG.frx":01C6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dtgLx"
      Tab(1).Control(1)=   "cmdPr(2)"
      Tab(1).Control(2)=   "cmdPr(1)"
      Tab(1).Control(3)=   "cmdPr(0)"
      Tab(1).ControlCount=   4
      Begin VB.Frame frmPx 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   465
         Left            =   0
         TabIndex        =   4
         Top             =   8730
         Width           =   9555
      End
      Begin VB.CommandButton cmdPr 
         Caption         =   "�ͻ�����"
         Height          =   345
         Index           =   0
         Left            =   -74940
         TabIndex        =   3
         Top             =   8790
         Width           =   1305
      End
      Begin VB.CommandButton cmdPr 
         Caption         =   "��  ��"
         Height          =   375
         Index           =   1
         Left            =   -73620
         TabIndex        =   2
         Top             =   8760
         Width           =   1185
      End
      Begin VB.CommandButton cmdPr 
         Caption         =   "��Ŀ����"
         Height          =   375
         Index           =   2
         Left            =   -72420
         TabIndex        =   1
         Top             =   8760
         Width           =   5295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgKh 
         Height          =   8475
         Left            =   0
         TabIndex        =   5
         Top             =   330
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   14949
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLx 
         Height          =   8475
         Left            =   -74970
         TabIndex        =   6
         Top             =   300
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   14949
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Label lblYwy 
      Caption         =   "Label6"
      Height          =   255
      Left            =   12150
      TabIndex        =   28
      Top             =   7680
      Width           =   765
   End
   Begin VB.Label lblFw 
      Height          =   285
      Left            =   11970
      TabIndex        =   24
      Top             =   5130
      Width           =   2475
   End
   Begin VB.Label Label4 
      Caption         =   "��Χ��"
      Height          =   315
      Left            =   11040
      TabIndex        =   19
      Top             =   3570
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "��ѯ������"
      Height          =   225
      Left            =   10710
      TabIndex        =   11
      Top             =   5670
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "ֵ��"
      Height          =   405
      Left            =   11250
      TabIndex        =   10
      Top             =   6300
      Width           =   495
   End
End
Attribute VB_Name = "frmKhbrG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XuanRen As Integer
Public adoKhBr As Object
Public adoRenBr As Object
Public adoYwy As Object
'


'

'


'
'Private Sub Form_Resize()
''cmdBack.Left = frmkhbrG.Width - cmdBack.Width - 500
''cmdBack.Top = frmkhbrG.Height - cmdBack.Height - 700
''dtgKh.Height = frmkhbrG.Height - 1300
''frmPx.Top = dtgKh.Height + 100
'End Sub
'
'Private Sub optA_Click()
'dtgKh.Col = 1
'dtgKh.Sort = 1
'End Sub
'
'Private Sub optB_Click()
'dtgKh.Col = 2
'dtgKh.Sort = 2
'End Sub
'
'Private Sub optC_Click()
'dtgKh.Col = 3
'dtgKh.Sort = 3
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'frmKhbrG.Visible = False
'frmZu.Enabled = True
'End Sub
'
'Private Sub txtZ_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim tt As String
'On Error Resume Next
'If KeyCode = 13 Then
'    If chkAll.Value = 0 Then
'    Select Case comLx.Text
'        Case "��Ŀ����"
'            tt = "khNewV_xmmc('" & comYwy.Text & "','" & txtZ.Text & "')"
'        Case "�ͻ�����"
'            tt = "khNewV_man('" & comYwy.Text & "','" & txtZ.Text & "')"
'        End Select
'    ElseIf chkAll.Value = 1 And mod1.KhK = 1 Then
'        Select Case comLx.Text
'        Case "��Ŀ����"
'            tt = "Select khman as �ͻ�����,khsex as �Ա�,xmmc as ��Ŀ����,kid,xid from vkhRenNew  where xmmc like '%" & txtZ.Text & "%' and bm='" & mod1.Bm & "'"
'        Case "�ͻ�����"
'            tt = "Select khman as �ͻ�����,khsex as �Ա�,xmmc as ��Ŀ����,kid,xid from vkhRenNew  where khman like '%" & txtZ.Text & "%' and bm='" & mod1.Bm & "'"
'        End Select
'    ElseIf chkAll.Value = 1 And mod1.KhK = 2 Then
'        Select Case comLx.Text
'        Case "��Ŀ����"
'            tt = "Select khman as �ͻ�����,khsex as �Ա�,xmmc as ��Ŀ����,kid,xid from vkhRenNew  where xmmc like '%" & txtZ.Text & "%'"
'        Case "�ͻ�����"
'            tt = "Select khman as �ͻ�����,khsex as �Ա�,xmmc as ��Ŀ����,kid,xid from vkhRenNew  where khman like '%" & txtZ.Text & "%'"
'        End Select
'    End If
'    frmKhbrG.adoRenBr.Close
'    frmKhbrG.adoRenBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'    Set frmKhbrG.dtgLx.DataSource = frmKhbrG.adoRenBr
'    tabCx.Tab = 1
'End If
'End Sub

Private Sub cmdBack_Click()
frmKhbrG.Visible = False
frmZu.Enabled = True
End Sub

Private Sub cmdFw_Click()
Set Ren.XForm = New frmKhbrG
Call mod1.RenXz("frmKhbrG", Me, 0)
Me.XuanRen = 1
End Sub

Private Sub cmdHg_Click()
Dim xmAdo As Object
Dim tt As String
Dim Xmmc As String
Dim Kid As Long
Dim xid As Long
On Error Resume Next
dtgKh.Col = 2
xid = Val(dtgKh.Text)
dtgKh.Col = 5
Kid = Val(dtgKh.Text)
dtgKh.Col = 1
Xmmc = dtgKh.Text



If lblYwy.Caption <> "" Then

    Set xmAdo = CreateObject("adodb.command")
    xmAdo.ActiveConnection = mod1.cc
    xmAdo.CommandText = "XMChange"
    xmAdo.CommandType = adCmdStoredProc
    xmAdo.Parameters("@ywy") = lblYwy.Caption
    xmAdo.Parameters("@uid") = lblYwy.ToolTipText
    xmAdo.Parameters("@xid") = xid
    xmAdo.Parameters("@xmmc") = Xmmc
    xmAdo.Parameters("@zf") = 1
    xmAdo.Parameters("@kk") = ""
    xmAdo.Execute
    If xmAdo.Parameters("@zf").Value = 0 Or IsNull(xmAdo.Parameters("@zf").Value) = True Then
        MsgBox "������ֹ���,������һ��,��������ύ���ɹ�,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        Exit Sub
    End If


    frmKhbrG.adoKhBr.Requery
    Set frmKhbrG.dtgKh.DataSource = frmKhbrG.adoKhBr
End If

End Sub

Private Sub cmdLz_Click()
Dim tt As String
On Error Resume Next

    tt = "Select * from Xmlz order by ҵ��Ա"
If mod1.DName = "֣��" Then
    tt = "Select * from Xmlz order by ҵ��Ա" And Qy = "�Ϻ�"
End If
    Set frmKhbrG.adoKhBr = CreateObject("adodb.recordset")
    frmKhbrG.adoKhBr.Close
    frmKhbrG.adoKhBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmKhbrG.dtgKh.DataSource = frmKhbrG.adoKhBr
    If frmKhbrG.adoKhBr.RecordCount > 0 Then
        frmKhbrG.dtgKh.FixedRows = 0
        frmKhbrG.dtgKh.MergeCol(4) = True
        frmKhbrG.dtgKh.MergeCol(12) = True
        frmKhbrG.dtgKh.MergeCol(14) = True
        frmKhbrG.dtgKh.MergeCells = 3
        frmKhbrG.dtgKh.FixedRows = 1
    End If
    tabCx.Tab = 0

End Sub

Private Sub cmdVall_Click()
Dim tt As String
On Error Resume Next
    If frmKhbrG.Visible = False Then Exit Sub
    If mod1.KhK = 1 Then
        tt = "Select * from XmView where ggl'" & mod1.DHid & "' order by ҵ��Ա"
    ElseIf mod1.KhK = 3 Then
        tt = "Select * from xmView order by comid,����,ҵ��Ա"
    ElseIf mod1.KhK = 2 And mod1.comId = 1 Then
        tt = "select * from xmView where comid=" & mod1.comId & " order by ����,ҵ��Ա"
    ElseIf mod1.KhK = 2 And mod1.comId = 0 Then '����
        tt = "select * from xmView where comid=" & mod1.comId & " and not(����='ά����3' or ����='��Ʒ��1' or ����='��Ʒ��2') order by ����,ҵ��Ա"
    End If
    Set frmKhbrG.adoKhBr = CreateObject("adodb.recordset")
    frmKhbrG.adoKhBr.Close
    frmKhbrG.adoKhBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmKhbrG.dtgKh.DataSource = frmKhbrG.adoKhBr
    If frmKhbrG.adoKhBr.RecordCount > 0 Then
        frmKhbrG.dtgKh.FixedRows = 0
        frmKhbrG.dtgKh.MergeCol(4) = True
        frmKhbrG.dtgKh.MergeCol(12) = True
        frmKhbrG.dtgKh.MergeCol(14) = True
        frmKhbrG.dtgKh.MergeCells = 3
        frmKhbrG.dtgKh.FixedRows = 1
    End If
    frmKhbrG.tabCx.Tab = 0
    frmKhbrG.Visible = True
End Sub

Private Sub Command1_Click()

End Sub


Private Sub cmdXQ_Click()
Dim tt As String
On Error Resume Next
Dim Kid As Long
Dim xid As Long

    dtgKh.Col = 2
    xid = Val(dtgKh.Text)
    dtgKh.Col = 5
    Kid = Val(dtgKh.Text)
    dtgKh.Col = 2
    If dtgKh.Text = "" Then Exit Sub
    frmWait.Show
    frmWait.ZOrder 0
    
    frmWait.Refresh
    frmWait.faWait.Play
    


    'wbDN.WindowState = 2

    'mod1.Kd = frmKhBrow.adoKh.Recordset.Fields("khdh").Value
    If mod1.DKZ(xid, 7) = True Then
        MsgBox "��ݱ�����" & mod1.DKRen & "��,���Ժ�����,������������ϵ."
        Exit Sub
    End If
    
    frmKhbrG.Enabled = False
    wbDN.Visible = False
    Me.MousePointer = 11
    '��¼����־
    Call mod1.zhuDa(7, dtgKh.Text)
    Call mod1.xmQing
    Call mod1.khQing
    Call mod1.xmBound(xid)
    wbDN.lblKid.Caption = wbDN.lblYZ.Tag
    Call mod1.khBound(wbDN.lblYZ.Tag, "yz")
    If Val(wbDN.lblXmPd.Caption) < 60 Then
        wbDN.frmJE.Visible = True
    End If
    wbDN.Left = 0
    wbDN.Top = 0
    wbDN.cmdMod.Enabled = True
    wbDN.cmdSave.Enabled = False
    Me.MousePointer = 0
    wbDN.tabKh.Tab = 0
    'wbDN.cmdRadd.Enabled = True
    If wbDN.txtKhmc.Text = "" Then
        wbDN.tabKh.TabEnabled(1) = False
    Else
        wbDN.tabKh.TabEnabled(1) = True
    End If
    wbDN.cmdSave.Enabled = True
    'wbDN.cmdSaveA.Enabled = True
    
    
    

    wbDN.modFi = False

    Me.MousePointer = 0
    wbDN.cmdSave.Enabled = False
    wbDN.tabKh.Enabled = True
    If wbDN.lblYwy.Caption = mod1.DName Or wbDN.lblXywy.Caption = mod1.DName Then
        wbDN.cmdMod.Enabled = True
    Else
        wbDN.cmdMod.Enabled = False
    End If
    wbDN.khAdd = False
    '����Ŀ��,Ĭ�ϵĴ򿪿ͻ�Ϊ��Ŀ����
    wbDN.optYz.Value = True
    wbDN.frmGL.Visible = False
    wbDN.frmJz.Visible = True
    frmWait.Visible = False
    wbDN.Visible = True
    wbDN.cmdMod.Enabled = True
    
    '���¶�̬ǩ�ְ�ť�ĳ�ʼ����
        For oo = 1 To 10
           wbDN.lblQM(oo).Left = wbDN.lblQM(oo - 1).Left + 1100
           wbDN.cmdQm(oo).Left = wbDN.cmdQm(oo - 1).Left + 1100
           wbDN.lblTm(oo).Left = wbDN.lblTm(oo - 1).Left + 1100
           mod1.HTP.MoveNext
        Next
End Sub

Private Sub cmdYwy_Click()
Set Ren.XForm = New frmKhbrG
Call mod1.RenXz("frmKhbrG", Me, 0)
Me.XuanRen = 2
End Sub

Private Sub dtgKH_DblClick()
Static Px As Boolean

If dtgKh.Row = 1 Then
    If Px = True Then
        dtgKh.Sort = 2
        Px = False
    Else
        dtgKh.Sort = 1
        Px = True
    End If
'Else
'    MsgBox MGa.ColData(1)
End If
End Sub


Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
dtgKh.ColWidth(0) = 300
dtgKh.ColWidth(1) = 4500
'dtgKH.ColWidth(4) = 700
dtgKh.ColWidth(5) = 0
dtgKh.ColWidth(6) = 0
dtgKh.ColWidth(7) = 0
dtgKh.ColWidth(8) = 0
dtgKh.ColWidth(9) = 0
dtgKh.ColWidth(10) = 0
dtgKh.ColWidth(11) = 0
dtgKh.ColWidth(13) = 0

dtgLx.ColWidth(0) = 300
dtgLx.ColWidth(3) = 5500
dtgLx.ColWidth(4) = 0
dtgLx.ColWidth(5) = 0
If mod1.DName = "������" Or mod1.DName = "֣��" Then
    cmdLZ.Visible = True
Else
    cmdLZ.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
frmKhbrG.Visible = False
frmZu.Enabled = True
Cancel = True
End If
End Sub


Private Sub NiceButton1_Click()
Dim tt As String
On Error Resume Next

comLx.Text = "��Ŀ����"


            'tt = "Select * from XmView where ��Ŀ���� like '%" & Trim(txtZ.Text) & "%' and comid=" & mod1.comId & "  and not(����='ά����3' or ����='��Ʒ��1' or ����='��Ʒ��2')��and lc=100 order by ҵ��Ա"
tt = "Select * from XmView where   lc=100 order by ҵ��Ա"

    Set frmKhbrG.adoKhBr = CreateObject("adodb.recordset")
    frmKhbrG.adoKhBr.Close
    frmKhbrG.adoKhBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmKhbrG.dtgKh.DataSource = frmKhbrG.adoKhBr
    If frmKhbrG.adoKhBr.RecordCount > 0 Then
        frmKhbrG.dtgKh.FixedRows = 0
        frmKhbrG.dtgKh.MergeCol(4) = True
        frmKhbrG.dtgKh.MergeCol(12) = True
        frmKhbrG.dtgKh.MergeCol(14) = True
        frmKhbrG.dtgKh.MergeCells = 3
        frmKhbrG.dtgKh.FixedRows = 1
    End If
    tabCx.Tab = 0

End Sub

Private Sub txtZ_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tt As String
On Error Resume Next
If KeyCode = 13 Then
comLx.Text = "��Ŀ����"
    Select Case comLx.Text
    Case "��Ŀ����"
        If mod1.KhK = 1 Then
            tt = "Select * from XmView where ��Ŀ���� like '%" & Trim(txtZ.Text) & "%'  and ����='" & mod1.Bm & "' order by ҵ��Ա"
        ElseIf mod1.KhK = 2 And mod1.comId <> 0 Then
            tt = "Select * from XmView where ��Ŀ���� like '%" & Trim(txtZ.Text) & "%' and comid=" & mod1.comId & " order by ҵ��Ա"
        ElseIf mod1.KhK = 3 And mod1.comId = 0 Or mod1.DName = "�Ǽ���" Or mod1.DName = "����" Then '����
            tt = "Select * from XmView where ��Ŀ���� like '%" & Trim(txtZ.Text) & "%' and  not(����='ά����3' or ����='��Ʒ��1' or ����='��Ʒ��2') order by ҵ��Ա"
        End If
    If mod1.Qy = "����" Then
            tt = "Select * from XmView where ��Ŀ���� like '%" & Trim(txtZ.Text) & "%' and (���� like '%����%' or ����='ά����ҵ������') order by ҵ��Ա"
    End If
    Case "�ͻ�����"
        tt = "khNewV_man('" & mod1.DName & "','" & txtZ.Text & "')"
    End Select
    Set frmKhbrG.adoKhBr = CreateObject("adodb.recordset")
    frmKhbrG.adoKhBr.Close
    frmKhbrG.adoKhBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmKhbrG.dtgKh.DataSource = frmKhbrG.adoKhBr
    If frmKhbrG.adoKhBr.RecordCount > 0 Then
        frmKhbrG.dtgKh.FixedRows = 0
        frmKhbrG.dtgKh.MergeCol(4) = True
        frmKhbrG.dtgKh.MergeCol(12) = True
        frmKhbrG.dtgKh.MergeCol(14) = True
        frmKhbrG.dtgKh.MergeCells = 3
        frmKhbrG.dtgKh.FixedRows = 1
    End If
    tabCx.Tab = 0
End If
End Sub


