VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmWork 
   BackColor       =   &H00C0FFC0&
   Caption         =   "ѯ�۵�"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15060
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   15060
   Begin VB.Frame frmGZ 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   9120
      TabIndex        =   39
      Top             =   8280
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CommandButton cmdGz 
         Caption         =   "����"
         Height          =   300
         Left            =   1560
         TabIndex        =   41
         Top             =   120
         Width           =   855
      End
      Begin MSComCtl2.DTPicker dtpGz 
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM"
         Format          =   111017987
         CurrentDate     =   41872
      End
   End
   Begin MSComCtl2.DTPicker dtpC 
      Height          =   255
      Left            =   2880
      TabIndex        =   38
      Top             =   8400
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Format          =   111017985
      CurrentDate     =   41843
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "��ѯ"
      Height          =   375
      Left            =   5760
      TabIndex        =   37
      Top             =   8400
      Width           =   975
   End
   Begin VB.TextBox txtZ 
      Height          =   270
      Left            =   2880
      TabIndex        =   36
      Top             =   8400
      Width           =   2535
   End
   Begin VB.ComboBox comCLx 
      Height          =   300
      ItemData        =   "frmWork.frx":0000
      Left            =   1200
      List            =   "frmWork.frx":000D
      TabIndex        =   35
      Text            =   "¼������"
      Top             =   8400
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgXZ 
      Height          =   5175
      Left            =   8520
      TabIndex        =   32
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12648384
      BackColorBkg    =   16777152
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   1815
      Left            =   9960
      TabIndex        =   31
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   3201
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgRen 
      Height          =   5175
      Left            =   6960
      TabIndex        =   30
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   12648384
      BackColorBkg    =   16777152
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   840
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2520
      Top             =   1800
   End
   Begin VB.Frame frmEdit 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�������༭"
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   6855
      Begin VB.CommandButton cmdCo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   765
         Left            =   4680
         Picture         =   "frmWork.frx":002B
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3240
         Width           =   675
      End
      Begin VB.TextBox txtRen 
         Height          =   270
         Left            =   4320
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtDate 
         Height          =   300
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CommandButton cmdCreate 
         BackColor       =   &H00FFFFC0&
         Caption         =   "���"
         Height          =   765
         Left            =   3240
         Picture         =   "frmWork.frx":046D
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3240
         Width           =   645
      End
      Begin VB.CommandButton cmdMod 
         BackColor       =   &H00FFFFC0&
         Caption         =   "�޸�"
         Height          =   765
         Left            =   3960
         Picture         =   "frmWork.frx":08AF
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "�޸�"
         Top             =   3240
         Width           =   675
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFC0&
         Caption         =   "����"
         Height          =   765
         Left            =   5390
         Picture         =   "frmWork.frx":0BB9
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "����"
         Top             =   3240
         Width           =   675
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ɾ��"
         Enabled         =   0   'False
         Height          =   765
         Left            =   6120
         Picture         =   "frmWork.frx":1223
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3240
         Width           =   675
      End
      Begin VB.TextBox txtZT 
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text6"
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txtLTime 
         Height          =   300
         Left            =   4320
         TabIndex        =   20
         Text            =   "Text5"
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtWtime 
         Height          =   300
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "Text4"
         Top             =   2760
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtpLt 
         Height          =   300
         Left            =   4320
         TabIndex        =   16
         Top             =   2160
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   529
         _Version        =   393216
         Format          =   111017986
         CurrentDate     =   41817
      End
      Begin MSComCtl2.DTPicker dtpFt 
         Height          =   300
         Left            =   1440
         TabIndex        =   14
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   111017986
         CurrentDate     =   41817
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   4320
         TabIndex        =   12
         Top             =   1560
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   529
         _Version        =   393216
         Format          =   111017985
         CurrentDate     =   41817
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   4320
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   1560
         Width           =   2415
      End
      Begin VB.ComboBox comLx 
         Height          =   300
         ItemData        =   "frmWork.frx":13AD
         Left            =   1440
         List            =   "frmWork.frx":13C9
         TabIndex        =   9
         Text            =   "ʩ��������"
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtXmmc 
         Height          =   270
         Left            =   4320
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtHid 
         Height          =   270
         Left            =   1440
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtDh 
         Height          =   270
         Left            =   1440
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ա"
         Height          =   255
         Left            =   3360
         TabIndex        =   28
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "�ܹ�ʱ"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "��;��ʱ"
         Height          =   255
         Left            =   3360
         TabIndex        =   19
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "������ʱ"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "����������"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ����"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "��ͬ���"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "���������"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   14208
      _Version        =   393216
      BackColor       =   16777152
      FixedCols       =   0
      BackColorFixed  =   15728356
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblGz 
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      Height          =   255
      Left            =   8040
      TabIndex        =   43
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label ud 
      BackStyle       =   0  'Transparent
      Caption         =   "�������ڣ�"
      Height          =   255
      Left            =   6960
      TabIndex        =   42
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "��ѯ����"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   8400
      Width           =   855
   End
End
Attribute VB_Name = "frmWork"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mid As Long
Dim timZm As Integer
Dim OT As Date

Private Sub cmdC_Click()
Dim tt As String
Select Case comCLx.Text
Case "¼������"
    tt = "select dh,'',xmmc,xz,cname,year(ft),month(ft),day(ft),cast(datepart(hour,ft) as nvarchar(2))+':'+cast(datepart(minute,ft) as nvarchar(2))," & _
        "cast(datepart(hour,lt) as nvarchar(2)) +':'+cast(datepart(minute,lt) as nvarchar(2)),lut,wt,zt,mid,ot from workDe where  year(edTime)=" & _
        Year(txtZ.Text) & " and month(edTime)=" & Month(txtZ.Text) & " and day(edTime)=" & Day(txtZ.Text) & " order by mid desc"
Case "����"
    tt = "select dh,'',xmmc,xz,cname,year(ft),month(ft),day(ft),cast(datepart(hour,ft) as nvarchar(2))+':'+cast(datepart(minute,ft) as nvarchar(2))," & _
        "cast(datepart(hour,lt) as nvarchar(2)) +':'+cast(datepart(minute,lt) as nvarchar(2)),lut,wt,zt,mid,ot from workDe" & _
        " where cname like '%" & txtZ.Text & "%'  order by mid desc"
Case "��Ŀ����"
    tt = "select dh,'',xmmc,xz,cname,year(ft),month(ft),day(ft),cast(datepart(hour,ft) as nvarchar(2))+':'+cast(datepart(minute,ft) as nvarchar(2))," & _
        "cast(datepart(hour,lt) as nvarchar(2)) +':'+cast(datepart(minute,lt) as nvarchar(2)),lut,wt,zt,mid,ot from workDe" & _
        " where xmmc like '%" & txtZ.Text & "%'  order by mid desc"
End Select
    Call Me.dtgBound(tt)
End Sub

Private Sub cmdCo_Click()
Mid = 0
txtRen.Text = ""
cmdSave.Enabled = True
End Sub

Private Sub cmdCreate_Click()
Call Me.Qing
cmdSave.Enabled = True
End Sub

Private Sub cmdDel_Click()
Dim ii As Integer
If Mid = 0 Then Exit Sub
ii = MsgBox("�Ƿ�ɾ���˼�¼?", vbQuestion + vbYesNo, "����")
If ii = vbNo Then Exit Sub


dtgRen.Visible = False

timZm = 2 'ɾ��
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "����ͳ��"
    mod1.cmd.Parameters("@NBLX") = "ɾ��"
    mod1.cmd.Parameters("@bh") = Mid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtXmmc.Text


    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtWtime.Text)

        mod1.cmd.Parameters("@mb1") = 0

    mod1.cmd.Parameters("@md1") = dtpFt.Value


    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        If timZm = 1 Then '����
            cmdSave.Enabled = False
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
End Sub

Private Sub cmdGz_Click()
Dim tt As String
Dim ii As Integer
Dim TDate As Date
Dim LDate As Date
ii = MsgBox("������" & Month(dtpGz.Value) & "�·ݽ��й��ʣ�����������¼��" & Month(dtpGz.Value) & "�·ݵĵ��ӣ����Զ���Ϊ" & Str((Month(dtpGz.Value) - 1)) & "�·ݽ���!", vbYesNo + vbInformation, "��ע��")
TDate = DateSerial(Year(dtpGz.Value), Month(dtpGz.Value) + 1, 1)
LDate = dtpGz.Value
If ii = vbNo Then Exit Sub
tt = "update workGz set Gdate='" & TDate & "'"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
dtpGz.Value = TDate
lblGz.Caption = Year(LDate) & "-" & Month(LDate)
End Sub

Private Sub cmdMod_Click()
If Year(dtpFt.Value) >= Year(dtpGz.Value) And Month(dtpFt.Value) >= Month(dtpGz.Value) Then
    If mod1.DName = "������" Or mod1.DName = "������" Then
        cmdSave.Enabled = True
        cmdDel.Enabled = True
    End If
Else
    MsgBox "�˵��Ѿ����ʣ������޸ģ�"
End If
End Sub

Private Sub cmdSave_Click()
If Me.txtDh.Text = "" Then
    MsgBox "�����빤������!"
    txtDh.SetFocus
    Exit Sub
End If
'''''If Me.txtHid.Text = "" Then
'''''    MsgBox "��������λ��ͬ���!"
'''''    txtHid.SetFocus
'''''    Exit Sub
'''''End If
If Me.txtXmmc.Text = "" Then
    MsgBox "��ȷ����Ŀ����!"
    txtHid.SetFocus
    Exit Sub
End If

dtgRen.Visible = False

timZm = 1 '����
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "����ͳ��"
    mod1.cmd.Parameters("@NBLX") = "����"
    mod1.cmd.Parameters("@bh") = Mid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtXmmc.Text
    mod1.cmd.Parameters("@mt2") = txtRen.Text
    mod1.cmd.Parameters("@mt4") = comLx.Text
    mod1.cmd.Parameters("@mt5") = txtHid.Text
    mod1.cmd.Parameters("@mt6") = txtDh.Text

    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtWtime.Text)
    mod1.cmd.Parameters("@mm2") = Val(txtLTime.Text)
    mod1.cmd.Parameters("@mm3") = Val(txtZT.Text)
    mod1.cmd.Parameters("@mm4") = Val(txtRen.ToolTipText)
        mod1.cmd.Parameters("@mb1") = 0

    mod1.cmd.Parameters("@md1") = dtpFt.Value
    mod1.cmd.Parameters("@md2") = dtpLt.Value
    mod1.cmd.Parameters("@md3") = OT
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "�ɹ�" Then
        MsgBox "������ֹ���,�����Źرճ���,��ִ�д˲���,�����Ȼʧ��,������������ϵ!"
        If timZm = 2 Then '����
            cmdSave.Enabled = False
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
End Sub

Private Sub comCLx_Change()
If comCLx.Text = "¼������" Then
    Me.dtpC.Visible = True
Else
    Me.dtpC.Visible = False
End If
End Sub

Private Sub Command1_Click()
Dim tt As String
Dim Ra
'�ȼ����������Ƿ���ȷ
tt = "select gdate from workgz"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
If Month(dtpGz.Value) <> Month(Ra(0, 0)) Then
    MsgBox "�����������ò���ȷ��"
    Exit Sub
End If

End Sub

Private Sub dtgBr_Click()
Dim tt As String
Dim Ra

dtgRen.Visible = False
dtgXZ.Visible = False
'If dtgRen.Text = "" Then Exit Sub
If dtgBr.Row < 1 Then Exit Sub
dtgN.Row = dtgBr.Row
dtgN.Col = 13
If Val(dtgN.Text) > 0 Then
 
    Mid = Val(dtgN.Text)
    tt = "select dh,hid,lut,zt,wt,xmmc,ft,lt,ft,cname,wid,ot from workDe where mid=" & Mid
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    On Error Resume Next
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    
    Me.txtDh.Text = Ra(0, 0)
    Me.txtHid.Text = Ra(1, 0)
    Me.txtLTime = Ra(2, 0)
    Me.txtZT.Text = Ra(3, 0)
    Me.txtWtime.Text = Ra(4, 0)
    Me.txtXmmc.Text = Ra(5, 0)

    Me.dtpFt.Value = Ra(6, 0)
    Me.dtpLt.Value = Ra(7, 0)
    Me.txtDate.Text = Ra(8, 0)
    Me.txtRen.Text = Ra(9, 0)
    Me.txtRen.ToolTipText = Ra(10, 0)
    OT = Ra(11, 0)
    frmEdit.Visible = True
End If

End Sub

Private Sub dtgRen_DblClick()
dtgRen.Col = 0
txtRen.Text = dtgRen.Text
dtgRen.Col = 1
txtRen.ToolTipText = dtgRen.Text
'dtgRen.Visible = False
End Sub


Private Sub dtgXZ_Click()
If dtgXZ.Left = txtXmmc.Left + txtXmmc.Width Then
    txtXmmc.Text = dtgXZ.Text
Else
    txtHid.Text = dtgXZ.Text
End If
dtgXZ.Visible = False
End Sub

Private Sub dtpC_CloseUp()
txtZ.Text = dtpC.Value
dtpC.Visible = False
txtZ.Visible = True
End Sub


Private Sub dtpDate_CloseUp()
txtDate.Text = dtpDate.Value
dtpFt.Value = dtpDate.Value & ":9.00"
dtpLt.Value = dtpDate.Value & ":17.30"
txtDate.Visible = True
dtpDate.Visible = False
End Sub


Private Sub dtpDate_LostFocus()
dtpDate.Visible = False
txtDate.Visible = True

End Sub


Private Sub dtpFt_Change()
Me.txtWtime = Round(DateDiff("n", dtpFt.Value, dtpLt.Value) / 60, 1)
Me.txtZT = Val(Me.txtWtime) / 60 + Val(Me.txtLTime)
If Month(dtpFt.Value) >= Month(dtpGz.Value) Then
    OT = dtpFt.Value
Else
    OT = DateSerial(Year(dtpGz.Value), Month(dtpGz.Value), 1)
End If
End Sub

Private Sub dtpLt_Change()
Me.txtWtime = Round(DateDiff("n", dtpFt.Value, dtpLt.Value) / 60, 1)
Me.txtZT = Val(Me.txtWtime) + Val(Me.txtLTime)
End Sub

Private Sub Form_Click()
dtgRen.Visible = False
dtgXZ.Visible = False
dtpC.Visible = False
txtZ.Visible = True
End Sub

Private Sub Form_DblClick()
If frmEdit.Visible = True Then
    frmEdit.Visible = False
Else
    frmEdit.Visible = True
End If
End Sub

Private Sub Form_Load()
Dim LDate As Date
Me.Width = mod1.FWidth + 500
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
Call dtgbrFF

Dim tt As String
Dim Ra
Dim La As Integer
Dim oo As Integer

Call Me.Qing

    tt = "select dh,'',xmmc,xz,cname,year(ft),month(ft),day(ft),cast(datepart(hour,ft) as nvarchar(2))+':'+cast(datepart(minute,ft) as nvarchar(2))," & _
        "cast(datepart(hour,lt) as nvarchar(2)) +':'+cast(datepart(minute,lt) as nvarchar(2)),lut,wt,zt,mid from workDe where not(dh is null) and edTime>='" & _
        DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda)) & "' order by mid desc"
    Call Me.dtgBound(tt)
    
    tt = "select gdate from workGz"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    On Error Resume Next
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    dtpGz.Value = Ra(0, 0)
    LDate = DateSerial(Year(Ra(0, 0)), Month(Ra(0, 0)) - 1, 1)
    lblGz.Caption = Year(LDate) & "-" & Month(LDate)
If mod1.DName <> "����ƽ" And mod1.DName <> "������" Then
    frmGZ.Visible = False
Else
    frmGZ.Visible = True
End If
End Sub
Private Sub dtgbrFF()
dtgBr.Clear
dtgN.Clear
    dtgBr.Cols = 15: dtgN.Cols = 15
    dtgBr.Row = 0
    dtgBr.Col = 0: dtgBr.Text = "���������": dtgBr.CellFontBold = True
    dtgBr.Col = 1: dtgBr.Text = "�ͻ����": dtgBr.CellFontBold = True
    dtgBr.Col = 2: dtgBr.Text = "��Ŀ����": dtgBr.CellFontBold = True
    dtgBr.Col = 3: dtgBr.Text = "��������":  dtgBr.CellFontBold = True
    dtgBr.Col = 4: dtgBr.Text = "ά����Ա": dtgBr.CellFontBold = True
    dtgBr.Col = 5: dtgBr.Text = "��": dtgBr.CellFontBold = True
    dtgBr.Col = 6: dtgBr.Text = "��": dtgBr.CellFontBold = True
    dtgBr.Col = 7: dtgBr.Text = "��": dtgBr.CellFontBold = True
    dtgBr.Col = 8: dtgBr.Text = "����ʱ��": dtgBr.CellFontBold = True
    dtgBr.Col = 9: dtgBr.Text = "���ʱ��": dtgBr.CellFontBold = True
    dtgBr.Col = 10: dtgBr.Text = "��;ʱ��": dtgBr.CellFontBold = True
    dtgBr.Col = 11: dtgBr.Text = "������ʱ": dtgBr.CellFontBold = True
    dtgBr.Col = 12: dtgBr.Text = "��ʱС��": dtgBr.CellFontBold = True
    dtgBr.Col = 13: dtgBr.Text = "mid": dtgBr.CellFontBold = True
    dtgBr.Col = 14: dtgBr.Text = "ot": dtgBr.CellFontBold = True
    dtgBr.ColWidth(0) = 1200
    dtgBr.ColWidth(2) = 3975
    dtgBr.ColWidth(5) = 735
    dtgBr.ColWidth(6) = 400
    dtgBr.ColWidth(7) = 400
    dtgBr.ColWidth(13) = 0
    dtgBr.ColWidth(1) = 0
    dtgBr.ColWidth(14) = 0
End Sub


Public Sub Qing()
Me.txtDh.Text = ""
Me.txtHid.Text = ""
Me.txtLTime = ""
Me.txtZT.Text = ""
Me.txtWtime.Text = ""
Me.txtXmmc.Text = ""
Me.dtpDate.Value = Date
Me.dtpFt.Value = Date & ":9.00"
Me.dtpLt.Value = Date & ":17.30"
Me.txtDate.Text = ""
Me.txtRen.Text = ""
Me.txtRen.ToolTipText = ""
OT = "1999-09-09"
Mid = 0
End Sub





Private Sub Form_Unload(Cancel As Integer)
frmZu.Enabled = True
End Sub

Private Sub timQuit_Timer()
Dim oo As Integer
Dim ii As Integer
Dim Rb, RC
Dim Qje As Single
Dim tt As String
On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0

If timZm = 1 Or timZm = 2 Then '���Ϊ��Ӻ�ͬ����
    cmdSave.Enabled = False
    cmdDel.Enabled = False
    tt = "select dh,'',xmmc,xz,cname,year(ft),month(ft),day(ft),cast(datepart(hour,ft) as nvarchar(2))+':'+cast(datepart(minute,ft) as nvarchar(2))," & _
        "cast(datepart(hour,lt) as nvarchar(2)) +':'+cast(datepart(minute,lt) as nvarchar(2)),lut,wt,zt,mid from workDe where not(dh is null) and edTime>='" & _
        DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda)) & "' order by mid desc"
    Call Me.dtgBound(tt)
    


ElseIf timZm = 3 Then
    Call Qing
End If
timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
Dim RC, RD, RE
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2,mt3 from ml where zid=" & mod1.Zid
Set mod1.WP = New ADODB.Recordset
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '�ύ�ɹ�
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdSave.Enabled = False
        Mid = mod1.WP.Fields("mm1").Value
    ElseIf timZm = 2 Then
        Mid = 0
    End If
    Call FmxcZcBr.Bound(FmxcZcBr.ETT)
    Exit Sub
    
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then 'δ���

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '����ʧ��
    timWait.Enabled = False
    ii = MsgBox("���������ڴ�����������ʱ,�������´���:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "��������!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdSave.Enabled = False
    End If
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("���������ڴ�����������ʱ,��ʱ!", vbExclamation + vbOKOnly, "��������!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdSave.Enabled = False
    End If
    Exit Sub

End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub


Private Sub txtDate_Click()
txtDate.Text = dtpDate.Value
txtDate.Visible = False
dtpDate.Visible = True

End Sub


Private Sub txtHid_Click()
Dim tt As String
Dim oo As Integer
Dim Ra
Dim La As Integer
dtgXZ.Left = txtHid.Left + txtHid.Width
If Len(txtXmmc.Text) < 3 Then Exit Sub
tt = "select hid from htping where xmmc ='" & txtXmmc.Text & "' and delf=1 and (htf=1 or htf=2 or htf=3)"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = Nothing
dtgXZ.Clear
Call Me.dtgXZFF
La = UBound(Ra, 2) + 1
dtgXZ.Rows = La + 50
For oo = 1 To La
    dtgXZ.Row = oo
    dtgXZ.Col = 0
    dtgXZ.Text = Ra(0, oo - 1)
'''    dtgXZ.Col = 1
'''    dtgXZ.Text = Ra(1, oo - 1)
Next
dtgXZ.Row = oo
dtgXZ.Text = "���޺�ͬ��"
dtgXZ.Visible = True
dtgXZ.TopRow = 1
dtgXZ.ColWidth(0) = 2000
End Sub

Private Sub txtHid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tt As String
Dim Ra
If KeyCode <> 13 Then Exit Sub
tt = "select xmmc from htping where hid=" & Val(txtHid.Text) & " and delf=1 "
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = Nothing
txtXmmc.Text = Ra(0, 0)


End Sub



Public Sub dtgRenFF()
dtgRen.Rows = 200
dtgRen.Row = 0
dtgRen.Text = "��ѡ��": dtgRen.CellFontBold = True
dtgRen.ColWidth(1) = 0
End Sub

Private Sub txtLTime_Change()
Me.txtWtime = Round(DateDiff("n", dtpFt.Value, dtpLt.Value) / 60, 1)
Me.txtZT = Val(Me.txtWtime) + Val(Me.txtLTime)
End Sub

Public Sub dtgBound(tt As String)
Dim Ra
Dim La As Long
Dim oo As Long
Call dtgbrFF
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgBr.Rows = La + 100
dtgN.Rows = La + 100
dtgBr.Visible = False
For oo = 1 To La
    dtgBr.Row = oo
    dtgBr.Col = 0: dtgBr.Text = Ra(0, oo - 1)
    dtgBr.Col = 1: dtgBr.Text = Ra(1, oo - 1)
    dtgBr.Col = 2: dtgBr.Text = Ra(2, oo - 1)
    dtgBr.Col = 3: dtgBr.Text = Ra(3, oo - 1)
    dtgBr.Col = 4: dtgBr.Text = Ra(4, oo - 1)
    dtgBr.Col = 5: dtgBr.Text = Ra(5, oo - 1)
    dtgBr.Col = 6: dtgBr.Text = Ra(6, oo - 1)
    dtgBr.Col = 7: dtgBr.Text = Ra(7, oo - 1)
    dtgBr.Col = 8: dtgBr.Text = Ra(8, oo - 1)
    dtgBr.Col = 9: dtgBr.Text = Ra(9, oo - 1)
    dtgBr.Col = 10: dtgBr.Text = Ra(10, oo - 1)
    dtgBr.Col = 11: dtgBr.Text = Ra(11, oo - 1)
    dtgBr.Col = 12: dtgBr.Text = Ra(12, oo - 1)
    dtgBr.Col = 13: dtgBr.Text = Ra(13, oo - 1)
    dtgBr.Col = 14: dtgBr.Text = Ra(14, oo - 1)
    dtgN.Row = oo
    dtgN.Col = 0: dtgN.Text = Ra(0, oo - 1)
    dtgN.Col = 1: dtgN.Text = Ra(1, oo - 1)
    dtgN.Col = 2: dtgN.Text = Ra(2, oo - 1)
    dtgN.Col = 3: dtgN.Text = Ra(3, oo - 1)
    dtgN.Col = 4: dtgN.Text = Ra(4, oo - 1)
    dtgN.Col = 5: dtgN.Text = Ra(5, oo - 1)
    dtgN.Col = 6: dtgN.Text = Ra(6, oo - 1)
    dtgN.Col = 7: dtgN.Text = Ra(7, oo - 1)
    dtgN.Col = 8: dtgN.Text = Ra(8, oo - 1)
    dtgN.Col = 9: dtgN.Text = Ra(9, oo - 1)
    dtgN.Col = 10: dtgN.Text = Ra(10, oo - 1)
    dtgN.Col = 11: dtgN.Text = Ra(11, oo - 1)
    dtgN.Col = 12: dtgN.Text = Ra(12, oo - 1)
    dtgN.Col = 13: dtgN.Text = Ra(13, oo - 1)
    dtgN.Col = 14: dtgN.Text = Ra(14, oo - 1)
Next
dtgBr.Visible = True
dtgBr.TopRow = 1
End Sub

Private Sub txtRen_Click()
dtgXZ.Visible = False
End Sub

Private Sub txtRen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 13 Then Exit Sub
Dim tt As String
Dim Ra
Dim La As Integer
Dim oo As Integer
If txtRen <> "" Then
    tt = "select Cname,uid from workWB where zzf=1 and cname like '%" & txtRen.Text & "%' order by cname"
Else
    tt = "select Cname,uid from workWB where zzf=1 order by cname"
End If
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = Nothing
Call Me.dtgRenFF
La = UBound(Ra, 2) + 1
For oo = 1 To La
    dtgRen.Row = oo
    dtgRen.Col = 0
    dtgRen.Text = Ra(0, oo - 1)
    dtgRen.Col = 1
    dtgRen.Text = Ra(1, oo - 1)
Next
dtgRen.Visible = True
End Sub

Private Sub txtXmmc_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tt As String
Dim oo As Integer
Dim Ra
Dim La As Integer
dtgXZ.Left = txtXmmc.Left + txtXmmc.Width
If Not (KeyCode = 13 And Len(txtXmmc.Text) >= 2) Then Exit Sub
tt = "select xmmc from htping where xmmc like '%" & txtXmmc.Text & "%' and delf=1 group by xmmc"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = Nothing
dtgXZ.Clear
Call Me.dtgXZFF
La = UBound(Ra, 2) + 1
dtgXZ.Rows = La + 50
For oo = 1 To La
    dtgXZ.Row = oo
    dtgXZ.Col = 0
    dtgXZ.Text = Ra(0, oo - 1)
'''    dtgXZ.Col = 1
'''    dtgXZ.Text = Ra(1, oo - 1)
Next
dtgXZ.Visible = True
dtgXZ.TopRow = 1
dtgXZ.ColWidth(0) = 3000
End Sub



Public Sub dtgXZFF()
dtgXZ.Rows = 200
dtgXZ.Row = 0
dtgXZ.Text = "��ѡ��": dtgXZ.CellFontBold = True
dtgXZ.ColWidth(1) = 0
dtgXZ.ColWidth(0) = 3000
End Sub

Private Sub txtZ_Click()
If comCLx.Text = "¼������" Then
    txtZ.Text = dtpC.Value
    txtZ.Visible = False
    dtpC.Visible = True
End If
End Sub

