VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmBB 
   Caption         =   "��������"
   ClientHeight    =   11370
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   14625
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11370
   ScaleWidth      =   14625
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdYjtj 
      Caption         =   "ҵ��ͳ��"
      Height          =   285
      Left            =   11550
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox comHtxz 
      Height          =   300
      ItemData        =   "frmBB.frx":0000
      Left            =   4740
      List            =   "frmBB.frx":0016
      TabIndex        =   15
      Text            =   "ȫ��"
      Top             =   420
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "�� ��"
      Height          =   285
      Left            =   990
      TabIndex        =   13
      Top             =   420
      Width           =   945
   End
   Begin VB.CommandButton cmdXuan 
      Caption         =   "ѡ ȡ"
      Height          =   285
      Left            =   0
      TabIndex        =   12
      Top             =   420
      Width           =   945
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "�򿪺�ͬ����"
      Height          =   285
      Left            =   11550
      TabIndex        =   11
      Top             =   60
      Width           =   1335
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "ȫ��"
      Height          =   285
      Left            =   10650
      TabIndex        =   10
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdFw 
      Caption         =   "ѡ��Ա������"
      Height          =   315
      Left            =   7860
      TabIndex        =   8
      Top             =   30
      Width           =   1425
   End
   Begin VB.CommandButton cmdCx 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��ѯ"
      Height          =   285
      Left            =   12990
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   90
      Width           =   795
   End
   Begin VB.Frame frmXz 
      Caption         =   "ѡ�����"
      Height          =   1365
      Left            =   0
      TabIndex        =   3
      Top             =   9270
      Visible         =   0   'False
      Width           =   14595
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBB 
      Height          =   8445
      Left            =   0
      TabIndex        =   2
      Top             =   750
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   14896
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "frmBB.frx":0044
      Left            =   900
      List            =   "frmBB.frx":0051
      TabIndex        =   1
      Text            =   "����ͳ�Ʊ�1"
      Top             =   30
      Width           =   3195
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   315
      Left            =   4740
      TabIndex        =   5
      Top             =   30
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   8454016
      Format          =   151781377
      CurrentDate     =   38797
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   315
      Left            =   6420
      TabIndex        =   6
      Top             =   30
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   8454016
      Format          =   151781377
      CurrentDate     =   38797
   End
   Begin VB.Label Label3 
      Caption         =   "��ͬ����:"
      Height          =   195
      Left            =   3870
      TabIndex        =   14
      Top             =   480
      Width           =   825
   End
   Begin VB.Label lblFw 
      Height          =   225
      Left            =   9360
      TabIndex        =   9
      Top             =   90
      Width           =   1185
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6090
      X2              =   6360
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Label Label2 
      Caption         =   "����:"
      Height          =   225
      Left            =   4230
      TabIndex        =   7
      Top             =   90
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "��������"
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   1155
   End
End
Attribute VB_Name = "frmBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoHT As ADODB.Recordset
Dim lb As String 'ѡ���ѯ����Ŀ
Dim LX As String '��Ӧ��ѯ��Ŀ�ļ�¼�������0

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText dtgBB.Clip
dtgBB.FixedRows = 1
End Sub

Private Sub cmdCx_Click()
Dim tt As String
Dim ii As Integer
Dim Fw As String '��Χ����
Dim RQ As String '��������
Dim FHg As Long
Dim Htxz As String '��ͬ����
On Error Resume Next
    
If comHtxz.Text = "ȫ��" Then
    Htxz = ""
ElseIf comHtxz.Text = "ά��" Then
    Htxz = " and (��ͬ����='ά��' or ��ͬ����='C. ά����ͬ')"
ElseIf comHtxz.Text = "����" Then
    Htxz = " and (��ͬ����='����' or ��ͬ����='D. ά�޺�ͬ')"
ElseIf comHtxz.Text = "�����" Then
    Htxz = " and (��ͬ����='�����' or ��ͬ����='A. �������ͬ')"
ElseIf comHtxz.Text = "��Ʒ" Then
    Htxz = " and (��ͬ����='��Ʒ' or ��ͬ����='E. ��Ʒ��ͬ')"
    ElseIf comHtxz.Text = "���̷ְ�" Then
    Htxz = " and ��ͬ����='���̷ְ�'"
End If
RQ = " where ��ͬ����>='" & dt1.Value & "' and ��ͬ����<'" & DateSerial(Year(dt2.Value), Month(dt2.Value), Day(dt2.Value) + 1) & "'"
'RQ = " where (��ͬ���� between '" & dt1.Value & "' and cast(cast('" & dt2.Value & "') as nvarchar(20) & ' 23:59:59.998') as smalldatetime))"
If lblFw.ToolTipText = "" Then
    If mod1.KhK = 1 And mod1.BM = "ά����1" Then
        Fw = " and ǩ������='" & mod1.BM & "' "
    ElseIf mod1.KhK = 1 And mod1.BM <> "ά����1" Then
        Fw = " and ǩ������='" & mod1.BM & "' "
    ElseIf mod1.KhK = 2 Then
        Fw = " and not(ǩ������='ά����3' or ǩ������='��Ʒ��1' or ǩ������='��Ʒ��2') and comid=" & mod1.comId & "  and ǩ������='" & lblFw.Caption & "' "
        If lblFw.Caption = "" Then
            Fw = " and not(ǩ������='ά����3' or ǩ������='��Ʒ��1' or ǩ������='��Ʒ��2') and comid=" & mod1.comId
        End If
    ElseIf mod1.KhK = 3 Then
        Fw = " and comid=" & mod1.comId & "  and ǩ������='" & lblFw.Caption & "' "
        If lblFw.Caption = "" Then
            Fw = " and comid=" & mod1.comId
        End If
    End If
Else
    Fw = " and ǩ����='" & lblFw.Caption & "'"
End If
Select Case comLx.Text
Case "����ͳ�Ʊ�1"
    dtgBB.ColWidth(0) = 300
    dtgBB.ColWidth(2) = 2000
    dtgBB.ColWidth(1) = 3000
    dtgBB.ColWidth(4) = 2000
    dtgBB.ColWidth(9) = 0
    dtgBB.ColWidth(8) = 0
    dtgBB.ColWidth(22) = 0
    tt = "select * from htyj" & RQ & Htxz & Fw & " order by ����,��������,ǩ����"
    adoHT.Close
    adoHT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    dtgBB.FixedRows = 1
    Set dtgBB.DataSource = adoHT
    FHg = 0
    
    dtgBB.Rows = adoHT.RecordCount + 2
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Col = 4
    dtgBB.Text = "�ϼ�"
    
    dtgBB.Col = 5
    FHg = 0
    For ii = 1 To adoHT.RecordCount
        dtgBB.Row = ii
        FHg = dtgBB.Text + FHg
    Next
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Text = FHg
Case "����ͳ�Ʊ�2"
    dtgBB.ColWidth(0) = 300
    dtgBB.ColWidth(2) = 0
    dtgBB.ColWidth(4) = 2000
    dtgBB.ColWidth(5) = 1500
    dtgBB.ColWidth(8) = 1000
    dtgBB.ColWidth(3) = 3000
    dtgBB.ColWidth(1) = 800
    tt = "select ǩ���� as ҵ��Ա,ǩ������ as ����,��Ŀ����,��ͬ���,��ͬ����,��ͬ���,��ͬ���� as ǩԼʱ��,��Ŀ���� as ����ë��,��ɱ��� from htyj" & RQ & Htxz & Fw & " order by ����,��������,ǩ����"
    adoHT.Close
    adoHT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    dtgBB.FixedRows = 1
    Set dtgBB.DataSource = adoHT
    '�����ͬ���
    FHg = 0
    dtgBB.Rows = adoHT.RecordCount + 2
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Col = 5
    dtgBB.Text = "�ϼ�"
    dtgBB.Col = 6
    FHg = 0
    For ii = 1 To adoHT.RecordCount
        dtgBB.Row = ii
        FHg = dtgBB.Text + FHg
    Next
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Text = FHg
    '��������
    FHg = 0
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Col = 7
    dtgBB.Text = "�ϼ�"
    dtgBB.Col = 8
    FHg = 0
    For ii = 1 To adoHT.RecordCount
        dtgBB.Row = ii
        FHg = dtgBB.Text + FHg
    Next
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Text = FHg
Case "����ͳ�Ʊ�3"
    dtgBB.ColWidth(0) = 300
    dtgBB.ColWidth(2) = 0
    dtgBB.ColWidth(4) = 2000
    dtgBB.ColWidth(5) = 1500
    dtgBB.ColWidth(8) = 1000
    dtgBB.ColWidth(3) = 3000
    dtgBB.ColWidth(1) = 800
    If lblFw.ToolTipText = "" Then
        If mod1.KhK = 1 Then
            Fw = " and ��������='" & mod1.BM & "' "

        ElseIf mod1.KhK = 2 Then
            Fw = " and not(ǩ������='ά����3' or ǩ������='��Ʒ��1' or ǩ������='��Ʒ��2') and comid=" & mod1.comId & " "
        ElseIf mod1.KhK = 3 Then
            Fw = " and comid=" & mod1.comId & " "
        End If
    Else
        Fw = " and ������='" & lblFw.Caption & "'"
    End If
    tt = "select ������ as ҵ��Ա,�������� as ����,��Ŀ����,��ͬ���,��ͬ����,��ͬ���,��ͬ���� as ǩԼʱ��,��Ŀ���� as ����ë��,��ɱ���,ǩ����,��Ŀ���� from htyj" & RQ & Htxz & Fw & " order by ����,��������,������,ǩ����"
    adoHT.Close
    adoHT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    dtgBB.FixedRows = 1
    Set dtgBB.DataSource = adoHT
    '�����ͬ���
    FHg = 0
    dtgBB.Rows = adoHT.RecordCount + 2
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Col = 5
    dtgBB.Text = "�ϼ�"
    dtgBB.Col = 6
    FHg = 0
    For ii = 1 To adoHT.RecordCount
        dtgBB.Row = ii
        FHg = dtgBB.Text + FHg
    Next
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Text = FHg
    '��������
    FHg = 0
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Col = 7
    dtgBB.Text = "�ϼ�"
    dtgBB.Col = 8
    FHg = 0
    For ii = 1 To adoHT.RecordCount
        dtgBB.Row = ii
        FHg = dtgBB.Text + FHg
    Next
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Text = FHg
    '������Ŀ����
    FHg = 0
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Col = 10
    dtgBB.Text = "�ϼ�"
    dtgBB.Col = 11
    FHg = 0
    For ii = 1 To adoHT.RecordCount
        dtgBB.Row = ii
        FHg = dtgBB.Text + FHg
    Next
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Text = FHg
End Select

'For oo = 7 To dtgFybb.Cols - 1
'    dtgFybb.Col = oo
'    FHg = 0
'    For ii = 1 To adoFyBB.RecordCount
'        dtgFybb.Row = ii
'        FHg = dtgFybb.Text + FHg
'    Next
'    dtgFybb.Row = dtgFybb.Rows - 1
'    dtgFybb.Text = FHg
'Next
'dtgFybb.Rows = dtgFybb.Rows + 1
'dtgFybb.Row = dtgFybb.Rows - 1
'dtgFybb.Col = 6
'dtgFybb.Text = "�ܼ�"
''�����ܼ�
'FHg = 0
'dtgFybb.Row = dtgFybb.Row - 1
'For oo = 7 To dtgFybb.Cols - 1
'    dtgFybb.Col = oo
'    FHg = FHg + dtgFybb.Text
'
'
'Next
'dtgFybb.Row = dtgFybb.Row + 1
'dtgFybb.Col = 7
'dtgFybb.Text = FHg
End Sub

Private Sub cmdFw_Click()
    Set Ren.XForm = New frmBB
    Call mod1.RenXz("frmBB", Me, 0)
End Sub


Private Sub cmdOpen_Click()
mod1.BTZ = 6
Dim tt As String
Dim xZ As String
Dim NewF As Boolean
Dim Hid As Long
'Dim Lid As String
On Error Resume Next
dtgBB.Col = 4

tt = "select htxz,hid,newF from htping where htbh='" & dtgBB.Text & "'"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
NewF = mod1.HTP.Fields("newF").Value
xZ = mod1.HTP.Fields("htxz").Value
Hid = mod1.HTP.Fields("hid").Value

'Lid = Str(Lid)
If mod1.DKZ(Hid, 1) = True Then
        MsgBox "��ݱ�����" & mod1.DKRen & "��,���Ժ�����,������������ϵ."
        Exit Sub
End If

frmWait.Visible = True
frmWait.ZOrder 0
frmWait.Refresh

If NewF = False Then
    If xZ = "C. ά����ͬ" Or xZ = "D. ά�޺�ͬ" Then
    wbHTP.Visible = False
    Call modHt.wbQing
    
    
    tt = "Select * from htping where hid=" & Hid
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Call modHt.wbBound
    
    
    '�򿪲��ϱ�
    tt = "Select * from htSale where htbh='" & wbHTP.txtHtbh.Text & "'"
    wbMx.adoRGF.Recordset.Close
    wbMx.adoRGF.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set wbMx.dtgSale.DataSource = wbMx.adoRGF
    wbMx.lblChg.Caption = wbHTP.txtClcb1.Text
    
    '��Ӧ�տ��
    tt = "Select * from htping1 where htBh='" & wbHTP.txtHtbh.Text & "' order by rq"
    frmFuK.adoHpt.Recordset.Close
    frmFuK.adoHpt.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set wbMx.dtgFk.DataSource = frmFuK.adoHpt
    
    '��Ӷ���
    tt = "Select * from Yongjin where htBh='" & wbHTP.txtHtbh.Text & "' order by yId"
    frmYj.adoYj.Recordset.Close
    frmYj.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set frmYj.dtgYj.DataSource = frmYj.adoYj
    

    wbHTP.Visible = True
    
    wbHTP.txtYj1.Visible = False
    wbHTP.txtYj2.Visible = False
    wbHTP.txtLr1.Visible = False
    wbHTP.txtLr2.Visible = False
    wbHTP.lblTcBe.Visible = False
    wbHTP.txtTcBe.Visible = False
    wbHTP.UpDa.Visible = False
    wbHTP.lblYj.Visible = False
    wbHTP.lblLr.Visible = False
    wbHTP.lblTC.Visible = False
    Exit Sub
    End If
    

    
    '������ͬ
    
    form2Htp.Visible = True
    mod1.workTt = ""
    mod1.workTt = "Select * from htPing where hid=" & Hid
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open mod1.workTt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    form2Htp.lblHtxz.Caption = ""
    
    Call modHt.htQing
    Call modHt.htBound '�󶨺�ͬ�����ֶ�
    

    '���տ��
    
    
    tt = "Select * from htPing1 where htBh='" & form2Htp.txtHtbh.Text & "' order by rq"
    frmFuK.adoHpt.Recordset.Close
    frmFuK.adoHpt.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    
    
    Set frmFuK.dtgFk.DataSource = frmFuK.adoHpt
    
    
    '�򿪲�Ʒ��
    tt = ""
    tt = "Select * from htSale Where htBh='" & form2Htp.txtHtbh.Text & "'"
    form2Htp.adoSale.Recordset.Close
    form2Htp.adoSale.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set form2Htp.dtgSale.DataSource = form2Htp.adoSale
    Set form2Htp.dtgYj.DataSource = form2Htp.adoSale
    Set form2Htp.dtgZj.DataSource = form2Htp.adoSale
    
     
    '��Ӷ���
    tt = "Select * from Yongjin where htBh='" & form2Htp.txtHtbh.Text & "' order by yId"
    frmYj.adoYj.Recordset.Close
    frmYj.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set frmYj.dtgYj.DataSource = frmYj.adoYj
    
    
    
    
    form2Htp.tabHt.TabEnabled(1) = True
    form2Htp.tabHt.TabEnabled(2) = True
    'End If
    
    
    
    
    
    
    
    form2Htp.tabHt.Tab = 0
    htBrow.MousePointer = 0
    
    
        'Ӷ������2����ɲ���ʾ
        form2Htp.txtYj1.Visible = False
        form2Htp.txtYj2.Visible = False
        form2Htp.txtLr1.Visible = False
        form2Htp.txtLr2.Visible = False
        'form2Htp.txtTc1.Visible = False
        'form2Htp.txtTc2.Visible = False
        form2Htp.lblYj.Visible = False
        form2Htp.lblLr2.Visible = False
        'form2Htp.lblTc.Visible = False
Else
        Call modHt.NewQing
        
        Call modHt.NewBound(Hid)

        frmWbNew.Visible = True

End If
End Sub

Private Sub cmdXuan_Click()
dtgBB.FixedRows = 0
End Sub

Private Sub cmdYjtj_Click()
Dim tt As String
Dim ii As Integer
Dim Fw As String '��Χ����
Dim RQ As String '��������
Dim FHg As Long
Dim Htxz As String '��ͬ����
On Error Resume Next
    
If comHtxz.Text = "ȫ��" Then
    Htxz = ""
ElseIf comHtxz.Text = "ά��" Then
    Htxz = " and (��ͬ����='ά��' or ��ͬ����='C. ά����ͬ')"
ElseIf comHtxz.Text = "����" Then
    Htxz = " and (��ͬ����='����' or ��ͬ����='D. ά�޺�ͬ')"
ElseIf comHtxz.Text = "�����" Then
    Htxz = " and (��ͬ����='�����' or ��ͬ����='A. �������ͬ')"
ElseIf comHtxz.Text = "��Ʒ" Then
    Htxz = " and (��ͬ����='��Ʒ' or ��ͬ����='E. ��Ʒ��ͬ')"
    ElseIf comHtxz.Text = "���̷ְ�" Then
    Htxz = " and ��ͬ����='���̷ְ�'"
End If
RQ = " where ��ͬ����>='" & dt1.Value & "' and ��ͬ����<'" & DateSerial(Year(dt2.Value), Month(dt2.Value), Day(dt2.Value) + 1) & "'"
If lblFw.ToolTipText = "" Then

        Fw = " and ǩ������='" & mod1.BM & "' "

Else
    Fw = " and ǩ����='" & lblFw.Caption & "'"
End If
    dtgBB.ColWidth(0) = 300
    dtgBB.ColWidth(2) = 0
    dtgBB.ColWidth(4) = 2000
    dtgBB.ColWidth(5) = 1500
    dtgBB.ColWidth(8) = 1000
    dtgBB.ColWidth(3) = 3000
    dtgBB.ColWidth(1) = 800
    If lblFw.ToolTipText = "" Then
        If mod1.KhK = 1 Then
            Fw = " and ǩ������='" & mod1.BM & "' "

        ElseIf mod1.KhK = 2 Then
            Fw = " and not(ǩ������='ά����3' or ǩ������='��Ʒ��1' or ǩ������='��Ʒ��2') and comid=" & mod1.comId & " "
        ElseIf mod1.KhK = 3 Then
            Fw = " and comid=" & mod1.comId & " "
        End If
    Else
        Fw = " and ������='" & lblFw.Caption & "'"
    End If
    tt = "select ������ as ҵ��Ա,�������� as ����,��Ŀ����,��ͬ���,��ͬ����,��ͬ���,��ͬ���� as ǩԼʱ��,��Ŀ���� as ����ë��,��ɱ���,ǩ���� from htyj" & RQ & Htxz & Fw & " order by ����,��������,������,ǩ����"
    adoHT.Close
    adoHT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    dtgBB.FixedRows = 1
    Set dtgBB.DataSource = adoHT
    '�����ͬ���
    FHg = 0
    dtgBB.Rows = adoHT.RecordCount + 2
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Col = 5
    dtgBB.Text = "�ϼ�"
    dtgBB.Col = 6
    FHg = 0
    For ii = 1 To adoHT.RecordCount
        dtgBB.Row = ii
        FHg = dtgBB.Text + FHg
    Next
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Text = FHg
    '��������
    FHg = 0
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Col = 7
    dtgBB.Text = "�ϼ�"
    dtgBB.Col = 8
    FHg = 0
    For ii = 1 To adoHT.RecordCount
        dtgBB.Row = ii
        FHg = dtgBB.Text + FHg
    Next
    dtgBB.Row = dtgBB.Rows - 1
    dtgBB.Text = FHg
End Sub

Private Sub comLx_Click()
If comLx.Text = "����ͳ�Ʊ�3" And mod1.DName = "Ф����" Then
    cmdYjtj.Visible = True
Else
    cmdYjtj.Visible = False
End If
End Sub

Private Sub Form_Load()
dtgBB.ColWidth(0) = 300
dtgBB.ColWidth(2) = 2000
dtgBB.ColWidth(1) = 3000
dtgBB.ColWidth(4) = 2000
dtgBB.ColWidth(9) = 0
dtgBB.ColWidth(8) = 0
dtgBB.ColWidth(22) = 0
Set adoHT = New ADODB.Recordset
dt1.Value = DateSerial(Year(Date), 1, 1)
dt2.Value = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmZu.WindowState = 0
End Sub

Private Sub Form_Resize()
dtgBB.Width = Me.Width
frmZu.WindowState = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmZu.WindowState = 0
End Sub


