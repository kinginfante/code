VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form FMXCXmmc 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ȷ����Ŀ���Ƽ��ͻ�����"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8595
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   8595
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3540
      Top             =   270
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5850
      Top             =   330
   End
   Begin NiceFormControl.NiceButton NiceButton1 
      Height          =   405
      Left            =   3660
      TabIndex        =   8
      Top             =   2400
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   714
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FMXCXmmc.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Style           =   5
      Caption         =   "�� �� �� ��"
   End
   Begin VB.ComboBox comKhmc 
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
      Height          =   360
      Left            =   3690
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1650
      Width           =   4425
   End
   Begin VB.TextBox txtXmmc 
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
      Height          =   360
      Left            =   3690
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   840
      Width           =   4395
   End
   Begin VB.TextBox txtX 
      BackColor       =   &H00FFFFC0&
      Height          =   300
      Left            =   2310
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   30
      Width           =   1125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   2475
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   4366
      _Version        =   393216
      BackColor       =   12648384
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   12648384
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "��������ѡ����Ӧ�Ŀͻ�"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   1350
      Width           =   2025
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "�������˫���б�ȷ����Ŀ����"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   30
      TabIndex        =   7
      Top             =   2910
      Width           =   3315
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "�ͻ�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4380
      TabIndex        =   6
      Top             =   1320
      Width           =   1725
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   4380
      TabIndex        =   5
      Top             =   540
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "����һ����������Ŀ�ؼ���"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   2205
   End
End
Attribute VB_Name = "FMXCXmmc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ka
Dim timZm As Integer
Dim Hid As Long
Dim Bid As Long
Public Lb As String '�����°�ѯ�۵����Ǻ�ͬ����

Private Sub comKhmc_Click()
On Error Resume Next
Dim tt As String
Dim Ra
Select Case comKhmc.ListIndex
Case 0
tt = "Select khdh from khzl where kid =" & Ka(7, 0)
Case 1
tt = "Select khdh from khzl where kid =" & Ka(8, 0)
Case 2
tt = "Select khdh from khzl where kid =" & Ka(9, 0)
Case 3
tt = "Select khdh from khzl where kid =" & Ka(10, 0)
Case 4
tt = "Select khdh from khzl where kid =" & Ka(11, 0)
Case 5
tt = "Select khdh from khzl where kid =" & Ka(12, 0)
Case 6
tt = "Select khdh from khzl where kid =" & Ka(13, 0)
End Select
If tt = "" Then Exit Sub
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
comKhmc.ToolTipText = Ra(0, 0)
End Sub

'
'
Private Sub dtgBr_DblClick()
Dim tt As String
Dim oo As Integer

Dim La
dtgBr.Col = 0
txtXmmc.Text = dtgBr.Text
dtgBr.Col = 1
txtXmmc.ToolTipText = dtgBr.Text

 tt = "Select yzmc,wymc,qt1mc,qt2mc,qt3mc,qt4mc,qt5mc,yzid,wyid,Qt1id,Qt2id,Qt3id,Qt4id,Qt5id from xmKhmc where xid=" & Val(txtXmmc.ToolTipText)
 Set mod1.HTP = CreateObject("adodb.recordset")
 mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
 If mod1.HTP.BOF = True Then
    Set mod1.HTP = Nothing
    Exit Sub
 End If
 Ka = mod1.HTP.GetRows
 mod1.HTP.Close
 Set mod1.HTP = Nothing
 La = UBound(Ka, 2) + 1

''''''    adoKhmc.Close
''''''    adoKhmc.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
On Error Resume Next
For oo = 6 To 0 Step -1
    comKhmc.RemoveItem oo
Next

    If La = 1 Then
        If IsNull(Ka(0, 0)) = False Then
            comKhmc.AddItem Ka(0, 0)
        Else
            comKhmc.AddItem " "
        End If
        If IsNull(Ka(1, 0)) = False Then
            comKhmc.AddItem Ka(1, 0)
        Else
            comKhmc.AddItem " "
        End If
        If IsNull(Ka(2, 0)) = False Then
            comKhmc.AddItem Ka(2, 0)
        Else
            comKhmc.AddItem " "
        End If
        If IsNull(Ka(3, 0)) = False Then
            comKhmc.AddItem Ka(3, 0)
        Else
            comKhmc.AddItem " "
        End If
        If IsNull(Ka(4, 0)) = False Then
            comKhmc.AddItem Ka(4, 0)
        Else
            comKhmc.AddItem " "
        End If
        If IsNull(Ka(5, 0)) = False Then
            comKhmc.AddItem Ka(5, 0)
        Else
            comKhmc.AddItem " "
        End If
        If IsNull(Ka(6, 0)) = False Then
            comKhmc.AddItem Ka(6, 0)
        Else
            comKhmc.AddItem " "
        End If
    End If

End Sub


Private Sub Form_Load()
dtgBr.ColWidth(1) = 0
Me.Height = 3645
Me.Width = 8685
dtgBr.ColWidth(0) = 3100
dtgBr.Rows = 500

End Sub

Public Sub Qing()
Dim oo As Integer
txtX.Text = ""
dtgBr.Clear
txtXmmc.Text = ""
txtXmmc.ToolTipText = ""
comKhmc.Text = ""
comKhmc.ToolTipText = ""
On Error Resume Next
For oo = 0 To 50
    comKhmc.RemoveItem (oo)
Next
End Sub

Private Sub NiceButton1_Click()
Dim tt As String
Dim Ra
If txtXmmc.Text = "" Or txtXmmc.ToolTipText = "" Then
    MsgBox ("��û��ѡ����ȷ����Ŀ!")
    Exit Sub
End If
If comKhmc.Text = "" Or comKhmc.ToolTipText = "" Then
    MsgBox ("��û��ѡ����ȷ�Ŀͻ�!")
    Exit Sub
End If

If Lb = "��ͬ����" Then
    If mod1.Qy = "�Ϻ�" Then
    '�ȼ�����Ŀ�Ƿ�ͨ�����
    tt = "select npf from xmzl where xid=" & Val(txtXmmc.ToolTipText)
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    If Not (Ra(0, 0) = True) Then
        MsgBox ("����Ŀ��δ�����г�Ӫ������ˣ�")
        Exit Sub
    End If
    End If
    timZm = 1
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.workKK
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�º�ͬ2013"
    mod1.cmd.Parameters("@NBLX") = "���"
    mod1.cmd.Parameters("@bh") = ""
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtXmmc.Text
    mod1.cmd.Parameters("@mt2") = comKhmc.Text
    mod1.cmd.Parameters("@mt3") = mod1.Qy
    mod1.cmd.Parameters("@mt4") = mod1.Bm
    mod1.cmd.Parameters("@mt5") = comKhmc.ToolTipText
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtXmmc.ToolTipText)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
   ' MsgBox "b"
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        If timZm = 1 Then
            cmdNew.Enabled = False
        End If
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If
    Set mod1.cmd = Nothing
Else
    If FmxcLxNew.LX = "" Or Val(FmxcLxNew.cmdNew.ToolTipText) = 0 Then
        MsgBox "��ѡ����ȷҵ������"
        Exit Sub
    End If
    timZm = 2
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.workKK
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "�º�ͬ2013"
    mod1.cmd.Parameters("@NBLX") = "���ѯ�۵�"
    mod1.cmd.Parameters("@bh") = ""
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtXmmc.Text
    mod1.cmd.Parameters("@mt2") = FmxcLxNew.LX 'ZL
    mod1.cmd.Parameters("@mt5") = comKhmc.ToolTipText
    mod1.cmd.Parameters("@mt25") = FmxcLxNew.Hid
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtXmmc.ToolTipText)
    mod1.cmd.Parameters("@mm2") = Val(FmxcLxNew.cmdNew.ToolTipText)
    FmxcLxNew.dtgNewLx.Col = 7
    mod1.cmd.Parameters("@mb1") = FmxcLxNew.dtgNewLx.Text
   ' Exit Sub
    mod1.cmd.Parameters("@md1") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
   ' MsgBox "b"
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        If timZm = 1 Then
            cmdNew.Enabled = False
        End If
        Exit Sub
    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If
    Set mod1.cmd = Nothing
End If

'�ɰ汾2012
''''''''''''If Lb = "��ͬ����" Then
''''''''''''    '�ȼ�����Ŀ�Ƿ�ͨ�����
''''''''''''    tt = "select npf from xmzl where xid=" & Val(txtXmmc.ToolTipText)
''''''''''''    Set mod1.HTP = CreateObject("adodb.recordset")
''''''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
''''''''''''    Ra = mod1.HTP.GetRows
''''''''''''    mod1.HTP.Close
''''''''''''    Set mod1.HTP = Nothing
''''''''''''    If Not (Ra(0, 0) = True) Then
''''''''''''        MsgBox ("����Ŀ��δ�����г�Ӫ������ˣ�")
''''''''''''        Exit Sub
''''''''''''    End If
''''''''''''
''''''''''''    timZm = 1
''''''''''''    Set mod1.cmd = CreateObject("adodb.command")
''''''''''''    mod1.cmd.ActiveConnection = mod1.workKK
''''''''''''    mod1.cmd.CommandText = "MLAdd"
''''''''''''    mod1.cmd.CommandType = adCmdStoredProc
''''''''''''    mod1.cmd.Parameters("@zid") = 0
''''''''''''    mod1.cmd.Parameters("@errch") = ""
''''''''''''    mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
''''''''''''    mod1.cmd.Parameters("@NBLX") = "���"
''''''''''''    mod1.cmd.Parameters("@bh") = ""
''''''''''''    mod1.cmd.Parameters("@ywy") = mod1.DName
''''''''''''    mod1.cmd.Parameters("@uid") = mod1.DHid
''''''''''''    mod1.cmd.Parameters("@mt1") = txtXmmc.Text
''''''''''''    mod1.cmd.Parameters("@mt2") = comKhmc.Text
''''''''''''    mod1.cmd.Parameters("@mt3") = mod1.Qy
''''''''''''    mod1.cmd.Parameters("@mt4") = mod1.Bm
''''''''''''    mod1.cmd.Parameters("@mt5") = comKhmc.ToolTipText
''''''''''''    mod1.cmd.Parameters("@mlt1") = ""
''''''''''''    mod1.cmd.Parameters("@mm1") = Val(txtXmmc.ToolTipText)
''''''''''''    mod1.cmd.Parameters("@mm2") = 0
''''''''''''    mod1.cmd.Parameters("@mb1") = 0
''''''''''''    mod1.cmd.Parameters("@md1") = Null
''''''''''''    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
''''''''''''    mod1.cmd.Execute
''''''''''''   ' MsgBox "b"
''''''''''''    mod1.Zid = mod1.cmd.Parameters("@zid").Value
''''''''''''    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
''''''''''''        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
''''''''''''        If timZm = 1 Then
''''''''''''            cmdNew.Enabled = False
''''''''''''        End If
''''''''''''        Exit Sub
''''''''''''    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
''''''''''''        Me.Enabled = False
''''''''''''        frmWaitA.Visible = True
''''''''''''        frmWaitA.Timer2.Enabled = False
''''''''''''
''''''''''''        frmWaitA.ZOrder 0
''''''''''''        frmWaitA.Timer2.Enabled = True
''''''''''''        timWait.Enabled = True
''''''''''''    End If
''''''''''''    Set mod1.cmd = Nothing
''''''''''''Else
''''''''''''
''''''''''''    timZm = 2
''''''''''''    Set mod1.cmd = CreateObject("adodb.command")
''''''''''''    mod1.cmd.ActiveConnection = mod1.workKK
''''''''''''    mod1.cmd.CommandText = "MLAdd"
''''''''''''    mod1.cmd.CommandType = adCmdStoredProc
''''''''''''    mod1.cmd.Parameters("@zid") = 0
''''''''''''    mod1.cmd.Parameters("@errch") = ""
''''''''''''    mod1.cmd.Parameters("@NB") = "�º�ͬ2011"
''''''''''''    mod1.cmd.Parameters("@NBLX") = "���ѯ�۵�"
''''''''''''    mod1.cmd.Parameters("@bh") = ""
''''''''''''    mod1.cmd.Parameters("@ywy") = mod1.DName
''''''''''''    mod1.cmd.Parameters("@uid") = mod1.DHid
''''''''''''    mod1.cmd.Parameters("@mt1") = txtXmmc.Text
''''''''''''    mod1.cmd.Parameters("@mt2") = FmxcLx.LX 'ZL
''''''''''''    mod1.cmd.Parameters("@mt5") = comKhmc.ToolTipText
''''''''''''    mod1.cmd.Parameters("@mlt1") = ""
''''''''''''    mod1.cmd.Parameters("@mm1") = Val(txtXmmc.ToolTipText)
''''''''''''    mod1.cmd.Parameters("@mm2") = Val(FmxcLx.cmdNew.ToolTipText)
''''''''''''    mod1.cmd.Parameters("@mb1") = 0
''''''''''''    mod1.cmd.Parameters("@md1") = Null
''''''''''''    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
''''''''''''    mod1.cmd.Execute
''''''''''''   ' MsgBox "b"
''''''''''''    mod1.Zid = mod1.cmd.Parameters("@zid").Value
''''''''''''    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
''''''''''''        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
''''''''''''        If timZm = 1 Then
''''''''''''            cmdNew.Enabled = False
''''''''''''        End If
''''''''''''        Exit Sub
''''''''''''    Else '�ύ�ɹ�,�ȴ�ϵͳ���Ĵ�������
''''''''''''        Me.Enabled = False
''''''''''''        frmWaitA.Visible = True
''''''''''''        frmWaitA.Timer2.Enabled = False
''''''''''''
''''''''''''        frmWaitA.ZOrder 0
''''''''''''        frmWaitA.Timer2.Enabled = True
''''''''''''        timWait.Enabled = True
''''''''''''    End If
''''''''''''    Set mod1.cmd = Nothing
''''''''''''End If

End Sub

Private Sub timQuit_Timer()
Dim htRow As Integer
Dim tt As String
Dim Rf
On Error Resume Next
Dim ii As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '���Ϊ��Ӻ�ͬ����
    Call FmxcNew.Bound(Hid)
    FmxcNew.Show
    FmxcNew.ZOrder 0
    FmxcNew.txtBz.Visible = False
    FmxcNew.cmdSave.Enabled = True
    FmxcNew.optXm.Visible = False
    FmxcNew.frmFk.Visible = True
    For ii = 0 To 4
        FmxcNew.Shape1(ii).Visible = True
    Next
    FmxcNew.comFPLX.Visible = True
    FmxcNew.companyId.Visible = True
    FmxcNew.dt3.Visible = True
    FmxcNew.dt4.Visible = True
    FmxcNew.txtXYwy.Locked = False
ElseIf timZm = 2 Then
    Call FmxcXJ.Bound(Bid)
    FmxcXJ.Show
    FmxcXJ.ZOrder 0
    FmxcXJ.cmdSave.Enabled = True
    
    
'�ɰ汾2012
'''''    HtRow = Val(FmxcLx.cmdNew.ToolTipText)\

'�°汾2013
htRow = Val(FmxcLxNew.cmdNew.ToolTipText)
    
'�°汾2013

    'If htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 And LX <> "����" Or htRow = 8 Or htRow >= 20 Then
    If (htRow = 1 Or htRow = 2 Or htRow = 3 Or htRow = 4 Or htRow = 5 Or htRow = 6 Or htRow = 7 And LX <> "����" Or htRow = 8 And FmxcLxNew.LX <> "����" Or htRow >= 20 Or InStr(1, "ά�����������˹�ѹ����ά�ޱ����н�ҵ��ְ��˷ѵ�װ�ѹ����˹�", FmxcLxNew.LX) > 0 Or _
     (Val(FmxcXJ.lblBid.ToolTipText) >= 20512 And FmxcXJ.lblZl.ToolTipText = True)) And Not (FmxcLxNew.LX = "�ְ�->�����˹�" And Val(FmxcXJ.lblBid.ToolTipText) > 22211 And Val(lblBid.ToolTipText) < 22670) Then

        FmxcXJ.frmWB.Visible = True
    Else
        FmxcXJ.frmSd.Visible = True
    End If
    If FmxcNew.Visible = True Then
        tt = "select zl,jhg,bianhao,0,lc,bid where htbh='" & Str(FmxcNew.lblHid.Caption) & "' order by bid"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
        Rf = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        Call FmxcNew.LXBound(Rf, Rg)
    End If
'�ɰ汾2012
'''''''    If HtRow = 1 Or HtRow = 2 Or HtRow = 3 Or HtRow = 4 Or HtRow = 6 Or HtRow = 12 Then
'''''''        FmxcXJ.frmWB.Visible = True
'''''''    Else
'''''''        FmxcXJ.frmSd.Visible = True
'''''''    End If
End If
timQuit.Enabled = False
Hid = 0
Me.Enabled = True
Me.Visible = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '�ύ�ɹ�
    mod1.Ti = 5
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    timWait.Enabled = False
    If timZm = 1 Then
        Hid = mod1.WP.Fields("bh").Value
    Else
        Bid = mod1.WP.Fields("bh").Value
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then 'δ���

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '����ʧ��
    ii = MsgBox("���������ڴ�����������ʱ,�������´���:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "��������!")
    Unload frmWaitA
    Me.Enabled = True
    If timZm = 1 Then
        NiceButton1.Enabled = False
    End If
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("���������ڴ�����������ʱ,��ʱ!", vbExclamation + vbOKOnly, "��������!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        NiceButton1.Enabled = False
    End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


Private Sub txtX_Change()
Dim oo As Integer
Dim tt As String
Dim Ra
Dim La
If Len(txtX.Text) > 1 Then
    tt = "select xmmc,xid from xmzl where xmmc like '%" & txtX.Text & "%' and uid='" & mod1.DHid & "'"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    If mod1.HTP.BOF = True Then
        Set mod1.HTP = Nothing
        Exit Sub
    End If
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2)
    For oo = 0 To La
        dtgBr.Row = oo
        dtgBr.Col = 0: dtgBr.Text = Ra(0, oo)
        dtgBr.Col = 1: dtgBr.Text = Ra(1, oo)
    Next
End If
End Sub

