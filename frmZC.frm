VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form fmxcZC 
   BackColor       =   &H00C0FFC0&
   Caption         =   "����"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15060
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   15060
   Begin VB.Frame frmGy 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��Ӧ��ѡ��"
      Height          =   2895
      Left            =   1920
      TabIndex        =   46
      Top             =   1080
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox txtGy 
         Height          =   315
         Left            =   0
         TabIndex        =   47
         Top             =   2160
         Width           =   4095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgGy 
         Height          =   1815
         Left            =   0
         TabIndex        =   48
         Top             =   240
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   3201
         _Version        =   393216
         BackColor       =   12648384
         Rows            =   50
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
   End
   Begin VB.CommandButton cmdBr 
      BackColor       =   &H00C0FFFF&
      Caption         =   "�顡��ѯ"
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7680
      Width           =   3975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   1935
      Left            =   12480
      TabIndex        =   42
      Top             =   2400
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   16777152
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
   Begin VB.CommandButton cmdTZ 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��ת"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox txtTid 
      Height          =   270
      Left            =   4320
      TabIndex        =   38
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton cmdCreate 
      BackColor       =   &H00C0FFC0&
      Caption         =   "���"
      Height          =   765
      Left            =   240
      Picture         =   "frmZC.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7440
      Width           =   645
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7095
      Left            =   6840
      TabIndex        =   22
      Top             =   0
      Width           =   7575
      Begin VB.CommandButton cmdD 
         BackColor       =   &H008080FF&
         Caption         =   "����"
         Height          =   345
         Left            =   1890
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   6720
         Width           =   1005
      End
      Begin VB.CommandButton cmdG 
         BackColor       =   &H00FF8080&
         Caption         =   "����"
         Height          =   345
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   6720
         Width           =   1005
      End
      Begin VB.CommandButton cmdA 
         BackColor       =   &H00FFFF00&
         Caption         =   "���"
         Height          =   345
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   6720
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker dtpRq 
         Height          =   255
         Left            =   2280
         TabIndex        =   30
         Top             =   5280
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         Format          =   101253121
         CurrentDate     =   41723
      End
      Begin VB.TextBox txtQje 
         Height          =   270
         Left            =   2520
         TabIndex        =   29
         Text            =   "Text3"
         Top             =   6240
         Width           =   3015
      End
      Begin VB.TextBox txtFph 
         Height          =   270
         Left            =   2520
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   5760
         Width           =   3015
      End
      Begin VB.TextBox txtQrQ 
         Height          =   270
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   5280
         Width           =   3015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
         Height          =   4695
         Left            =   0
         TabIndex        =   23
         Top             =   120
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   8281
         _Version        =   393216
         BackColor       =   16777152
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
      Begin VB.Label lblQid 
         Caption         =   "Label12"
         Height          =   255
         Left            =   6120
         TabIndex        =   35
         Top             =   6240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmZC.frx":0442
         Height          =   255
         Left            =   840
         TabIndex        =   26
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmZC.frx":0452
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   5280
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ɾ��"
      Enabled         =   0   'False
      Height          =   765
      Left            =   2400
      Picture         =   "frmZC.frx":045E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7440
      Width           =   675
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "����"
      Height          =   765
      Left            =   1680
      Picture         =   "frmZC.frx":05E8
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "����"
      Top             =   7440
      Width           =   675
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H00C0FFC0&
      Caption         =   "�޸�"
      Height          =   765
      Left            =   960
      Picture         =   "frmZC.frx":0C52
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "�޸�"
      Top             =   7440
      Width           =   675
   End
   Begin VB.Frame frmYY 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   5040
      Width           =   6495
      Begin VB.Timer timQuit 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.TextBox txtYY 
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Text            =   "frmZC.frx":0F5C
         Top             =   480
         Width           =   5895
      End
      Begin VB.Label lblZCid 
         Caption         =   "Label12"
         Height          =   255
         Left            =   5280
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊδ��ԭ��"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame FmxcNew 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "׷�ӳɱ���"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.ComboBox companyId 
         Height          =   300
         ItemData        =   "frmZC.frx":0F62
         Left            =   1800
         List            =   "frmZC.frx":0F72
         TabIndex        =   45
         Text            =   "�Ϻ���������յ��������޹�˾"
         Top             =   120
         Width           =   4215
      End
      Begin VB.TextBox txtWCF 
         Height          =   270
         Left            =   1800
         TabIndex        =   41
         Text            =   "Text1"
         Top             =   4440
         Width           =   4215
      End
      Begin VB.Timer timWait 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   480
      End
      Begin VB.TextBox txtZrq 
         Height          =   270
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   3902
         Width           =   4215
      End
      Begin VB.TextBox txtCg 
         Height          =   270
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text6"
         Top             =   3366
         Width           =   4215
      End
      Begin VB.TextBox txtFje 
         Height          =   270
         Left            =   1800
         TabIndex        =   11
         Text            =   "Text5"
         Top             =   2830
         Width           =   4215
      End
      Begin VB.TextBox txtFrq 
         Height          =   270
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text4"
         Top             =   2294
         Width           =   4215
      End
      Begin VB.TextBox txtHtbh 
         Height          =   270
         Left            =   1800
         TabIndex        =   9
         Text            =   "Text3"
         Top             =   1758
         Width           =   4215
      End
      Begin VB.TextBox txtYhh 
         Height          =   270
         Left            =   1800
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   1222
         Width           =   4215
      End
      Begin VB.TextBox txtGymc 
         Height          =   270
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   686
         Width           =   4215
      End
      Begin MSComCtl2.DTPicker dtpFk 
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   2160
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         _Version        =   393216
         Format          =   101253121
         CurrentDate     =   41723
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "��˾����"
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "��ɷ�"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmZC.frx":0FE2
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   3900
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ӧ������"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmZC.frx":0FF6
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmZC.frx":1006
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmZC.frx":1012
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmZC.frx":1020
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   2820
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmZC.frx":102E
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   3360
         Width           =   975
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "���:"
      Height          =   255
      Left            =   3600
      TabIndex        =   37
      Top             =   7800
      Width           =   615
   End
End
Attribute VB_Name = "fmxcZC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer

Private Sub cmdA_Click()
If txtQrQ.Text = "" Then
    MsgBox "����������!"
    txtQrQ.Visible = True
    txtQrQ.SetFocus
    Exit Sub
End If
If txtFph.Text = "" Then
    MsgBox "�����뷢Ʊ��!"
    txtFph.SetFocus
    Exit Sub
End If
If txtQje.Text = "" Then
    MsgBox "����������֧���!"
    txtQje.SetFocus
    Exit Sub
End If

timZm = 2 '���۱༭
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��Ʊ��֧"
    mod1.cmd.Parameters("@NBLX") = "��ϸ�༭"
    mod1.cmd.Parameters("@bh") = ""
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "���"
    mod1.cmd.Parameters("@mt2") = txtFph.Text
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtQje.Text)
    mod1.cmd.Parameters("@mm2") = Val(lblZCid.Caption)

        mod1.cmd.Parameters("@mb1") = 0

    mod1.cmd.Parameters("@md1") = txtQrQ.Text


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

Private Sub cmdBr_Click()
FmxcZcBr.Show 0
End Sub

Private Sub cmdCreate_Click()
Call Qing

End Sub

Private Sub cmdD_Click()
Dim ii As Integer
If Val(lblQid.Caption) = 0 Then Exit Sub
ii = MsgBox("�Ƿ�ɾ����һ��?", vbYesNo + vbQuestion, "��ȷ��")
If ii = vbNo Then Exit Sub

timZm = 2 '���۱༭
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��Ʊ��֧"
    mod1.cmd.Parameters("@NBLX") = "��ϸ�༭"
    mod1.cmd.Parameters("@bh") = lblQid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "ɾ��"
    mod1.cmd.Parameters("@mt2") = txtFph.Text
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtQje.Text)
    mod1.cmd.Parameters("@mm2") = 0

        mod1.cmd.Parameters("@mb1") = 0

    mod1.cmd.Parameters("@md1") = txtQrQ.Text


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

Private Sub cmdDel_Click()
Dim ii As Integer
ii = MsgBox("�Ƿ�ɾ��������¼?", vbYesNo + vbQuestion, "��ȷ��")
If ii = vbNo Then Exit Sub
timZm = 3 '����
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��Ʊ��֧"
    mod1.cmd.Parameters("@NBLX") = "ɾ��"
    mod1.cmd.Parameters("@bh") = Trim(lblZCid.Caption)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtGymc.Text
    mod1.cmd.Parameters("@mt2") = txtGymc.ToolTipText
    mod1.cmd.Parameters("@mt3") = txtYhh.Text
    mod1.cmd.Parameters("@mt4") = txtHtbh.Text
    mod1.cmd.Parameters("@mt5") = txtCg.Text
    mod1.cmd.Parameters("@mt6") = txtCg.ToolTipText
    mod1.cmd.Parameters("@mlt1") = txtYY.Text
    mod1.cmd.Parameters("@mm1") = Val(txtFje.Text)
    mod1.cmd.Parameters("@mm2") = 0

        mod1.cmd.Parameters("@mb1") = 0 'ȫ����

    mod1.cmd.Parameters("@md1") = txtFrq.Text
    mod1.cmd.Parameters("@md2") = txtZrq.Text

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

Private Sub cmdG_Click()
If Val(lblQid.Caption) = 0 Then Exit Sub
If txtQrQ.Text = "" Then
    MsgBox "����������!"
    txtQrQ.SetFocus
    Exit Sub
End If
If txtFph.Text = "" Then
    MsgBox "�����뷢Ʊ��!"
    txtFph.SetFocus
    Exit Sub
End If
If txtQje.Text = "" Then
    MsgBox "����������֧���!"
    txtQje.SetFocus
    Exit Sub
End If

timZm = 2 '���۱༭
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��Ʊ��֧"
    mod1.cmd.Parameters("@NBLX") = "��ϸ�༭"
    mod1.cmd.Parameters("@bh") = lblQid.Caption
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = "����"
    mod1.cmd.Parameters("@mt2") = txtFph.Text
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtQje.Text)
    mod1.cmd.Parameters("@mm2") = Val(lblZCid.Caption)

        mod1.cmd.Parameters("@mb1") = 0

    mod1.cmd.Parameters("@md1") = txtQrQ.Text


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

Private Sub cmdMod_Click()
If mod1.DName = "��Ƽ" Or mod1.DName = "�˫�" Or mod1.DName = "������" Then
    cmdSave.Enabled = True
    cmdDel.Enabled = True
End If
End Sub

Private Sub cmdSave_Click()
If txtGymc.Text = "" Then
    MsgBox "�����빩Ӧ������!"
    txtGymc.SetFocus
    Exit Sub
End If
If txtYhh.Text = "" Then
    MsgBox "������������ˮ��!"
    txtYhh.SetFocus
    Exit Sub
End If
If txtHtbh.Text = "" Then
    MsgBox "�������ͬ��!"
    txtHtbh.SetFocus
    Exit Sub
End If
If txtFrq.Text = "" Then
    MsgBox "�����븶������!"
    txtFrq.SetFocus
    Exit Sub
End If
If txtFje.Text = "" Then
    MsgBox "�����븶����!"
    txtFje.SetFocus
    Exit Sub
End If
If txtCg.Text = "" Then
    MsgBox "������ɹ�Ա!"
    txtCg.SetFocus
    Exit Sub
End If
timZm = 1 '����
    Set mod1.cmd = New ADODB.command
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "��Ʊ��֧"
    mod1.cmd.Parameters("@NBLX") = "����"
    mod1.cmd.Parameters("@bh") = Trim(lblZCid.Caption)
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtGymc.Text
    mod1.cmd.Parameters("@mt2") = txtGymc.ToolTipText
    mod1.cmd.Parameters("@mt3") = txtYhh.Text
    mod1.cmd.Parameters("@mt4") = txtHtbh.Text
    mod1.cmd.Parameters("@mt5") = txtCg.Text
    mod1.cmd.Parameters("@mt6") = txtCg.ToolTipText
    mod1.cmd.Parameters("@mt7") = companyId.Text
    mod1.cmd.Parameters("@mlt1") = txtYY.Text
    mod1.cmd.Parameters("@mm1") = Val(txtFje.Text)
    mod1.cmd.Parameters("@mm2") = 0

        mod1.cmd.Parameters("@mb1") = 0 'ȫ����

    mod1.cmd.Parameters("@md1") = txtFrq.Text
    mod1.cmd.Parameters("@md2") = txtZrq.Text

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

Private Sub cmdTZ_Click()
Call Me.Bound(Val(txtTid.Text))
End Sub

Private Sub dtgBr_Click()
On Error Resume Next
dtgN.Row = dtgBr.Row
dtgN.Col = 1: dtpRq.Value = dtgN.Text: txtQrQ.Text = dtgN.Text
dtgN.Col = 2: txtFph.Text = dtgN.Text
dtgN.Col = 3: txtQje.Text = dtgN.Text
dtgN.Col = 4: lblQid.Caption = dtgN.Text
End Sub

Private Sub dtgGy_DblClick()
dtgGy.Col = 0: txtGymc.Text = dtgGy.Text
dtgGy.Col = 1: txtGymc.ToolTipText = dtgGy.Text
frmGy.Visible = False
End Sub


Private Sub dtpFk_CloseUp()
dtpFk.Visible = False
txtFrq.Text = DateSerial(Year(dtpFk.Value), Month(dtpFk.Value), Day(dtpFk.Value))
txtZrq.Text = DateSerial(Year(txtFrq.Text), Month(txtFrq.Text), Day(dtpFk.Value) + 7)
txtFrq.Visible = True
End Sub



Public Sub Qing()
Me.txtCg.Text = "": Me.txtCg.ToolTipText = ""
Me.txtFje.Text = ""
Me.txtFrq.Text = ""
Me.txtGymc.Text = "": Me.txtGymc.ToolTipText = ""
Me.txtHtbh.Text = ""
Me.txtYhh.Text = ""
Me.dtpFk.Value = mod1.DQda
Me.dtpFk.Left = Me.txtFrq.Left
Me.dtpFk.Visible = False
Me.txtZrq.Text = ""
Me.txtYY.Text = ""
Me.txtWCF.Text = "δ���"

Me.txtQje.Text = ""
Me.txtFph.Text = ""
Me.txtQrQ.Text = ""
Me.dtpRq.Value = mod1.DQda
Me.dtpRq.Left = Me.txtQrQ.Left
Me.cmdSave.Enabled = False
Me.cmdDel.Enabled = False

Me.lblZCid.Caption = ""
Me.dtpRq.Visible = False
Me.txtQrQ.Visible = True
Me.companyId.Text = ""
Call Me.dtgbrFF
End Sub



Private Sub dtPRQ_CloseUp()
txtQrQ.Text = DateSerial(Year(dtpRq.Value), Month(dtpRq.Value), Day(dtpRq.Value))
dtpRq.Visible = False
txtQrQ.Visible = True
End Sub


Private Sub Form_DblClick()
frmGy.Visible = False
End Sub

Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmZu.Enabled = True
End Sub

Private Sub timQuit_Timer()
Dim oo As Integer
Dim ii As Integer
Dim Rb, RC
Dim Qje As Single

On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
Dim tt As String
If timZm = 1 Then '���Ϊ��Ӻ�ͬ����
    cmdSave.Enabled = False
ElseIf timZm = 2 Then
tt = "select qrq,fph,qje,qid from zcbDetail where zcid=" & Val(lblZCid.Caption) & " and delf=1 order by qid;" & _
    "select sum(qje) from zcbDetail where zcid=" & Val(lblZCid.Caption) & " and delf=1"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    On Error Resume Next
    Rb = mod1.HTP.GetRows
    Set mod1.HTP = mod1.HTP.NextRecordset
    RC = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
        
      
    If IsNull(RC(0, 0)) = False Then Qje = RC(0, 0)
    If Qje >= Val(Me.txtFje.Text) Then
        txtWCF.Text = "���"
    Else
        txtWCF.Text = "δ���"
    End If
    Call Me.dtgBound(Rb)

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
        lblZCid.Caption = mod1.WP.Fields("mm1").Value
        txtTid.Text = mod1.WP.Fields("mm1").Value
        
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

Private Sub txtCg_DblClick()

If txtCg.Text = "������" Then
    txtCg.Text = "�����"
    txtCg.ToolTipText = "HM804"
ElseIf txtCg.Text = "�����" Then
    txtCg.Text = "�޳�"
    txtCg.ToolTipText = "HM219"
ElseIf txtCg.Text = "�޳�" Then
    txtCg.Text = "������"
    txtCg.ToolTipText = "HM639"
ElseIf txtCg.Text = "������" Then
    txtCg.Text = "�����"
    txtCg.ToolTipText = "HM794"
ElseIf txtCg.Text = "�����" Or txtCg.Text = "" Then
    txtCg.Text = "������"
    txtCg.ToolTipText = "HM651"

End If
End Sub


Private Sub txtFrq_Click()
If cmdSave.Enabled = False Then Exit Sub
dtpFk.Visible = True
txtFrq.Visible = False
txtFrq.Text = dtpFk.Value
End Sub

Public Sub dtgbrFF()
dtgBr.Clear
dtgBr.Cols = 5
dtgBr.Rows = 50
dtgBr.Row = 0
dtgBr.Col = 1: dtgBr.Text = "����": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "��Ʊ��": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "����֧���": dtgBr.CellFontBold = True
dtgBr.ColWidth(4) = 0
dtgBr.ColWidth(0) = 345
dtgBr.ColWidth(1) = 2100
dtgBr.ColWidth(2) = 3045
dtgBr.ColWidth(3) = 1635
dtgN.Clear
dtgN.Cols = 5
dtgN.Rows = 50
dtgN.Row = 0
dtgN.Col = 1: dtgN.Text = "����": dtgN.CellFontBold = True
dtgN.Col = 2: dtgN.Text = "��Ʊ��": dtgN.CellFontBold = True
dtgN.Col = 3: dtgN.Text = "����֧���": dtgN.CellFontBold = True
dtgN.ColWidth(4) = 0

End Sub

Private Sub txtGy_Change()
Dim tt As String
Dim Ra
Dim La As Long
Dim oo As Long
If Len(txtGy.Text) < 2 Then Exit Sub
'tt = "select mc,gid from gymxc where mc like '%" & txtGy.Text & "%' and delf=1 and lc=100"
tt = "select mc,gid from gymxc where mc like '%" & txtGy.Text & "%' and delf=1 and lc>=2"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Call Me.dtgGYFF
For oo = 1 To La
    dtgGy.Row = oo
    dtgGy.Col = 0: dtgGy.Text = Ra(0, oo - 1)
    dtgGy.Col = 1: dtgGy.Text = Ra(1, oo - 1)
Next
End Sub

Public Sub dtgGYFF()
dtgGy.Clear
dtgGy.Rows = 50
dtgGy.Cols = 2
dtgGy.Row = 0
dtgGy.Col = 0: dtgGy.Text = "��Ӧ�����ƣ����˫��ѡ��": dtgGy.CellFontBold = True
dtgGy.ColWidth(1) = 0
dtgGy.ColWidth(0) = 3480

End Sub
Private Sub txtGymc_DblClick()
frmGy.Visible = True
End Sub

Private Sub txtQrQ_Click()
dtpRq.Visible = True
txtQrQ.Visible = False
txtQrQ.Text = DateSerial(Year(dtpRq.Value), Month(dtpRq.Value), Day(dtpRq.Value))
End Sub



Public Sub Bound(ZCid As Long)
Dim tt As String
Dim Ra
Dim Rb
Dim RC
Dim Qje As Single
Call Me.Qing
tt = "select * from zcb where zcid=" & ZCid & " and delf=1;" & _
    "select qrq,fph,qje,qid from zcbDetail where zcid=" & ZCid & " and delf=1 order by qid;" & _
    "select sum(qje) from zcbDetail where zcid=" & ZCid & " and delf=1"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
    
Me.txtGymc.Text = Ra(0, 0)
Me.txtGymc.ToolTipText = Ra(1, 0)
Me.txtYhh.Text = Ra(2, 0)
Me.txtHtbh.Text = Ra(3, 0)
Me.txtFrq.Text = Ra(4, 0)
Me.txtFje.Text = Ra(5, 0)
Me.txtCg.Text = Ra(6, 0)
Me.txtCg.ToolTipText = Ra(7, 0)
Me.txtZrq.Text = Ra(8, 0)
Me.txtYY.Text = Ra(9, 0)
Me.lblZCid.Caption = Ra(10, 0)
Me.companyId.Text = Ra(11, 0)
    
If IsNull(RC(0, 0)) = False Then Qje = RC(0, 0)
If Qje >= Val(Me.txtFje.Text) Then
    txtWCF.Text = "���"
Else
    txtWCF.Text = "δ���"
End If
Call Me.dtgBound(Rb)
End Sub

Public Sub dtgBound(Rb)
Dim Lb As Integer
Dim oo As Integer
On Error Resume Next
dtgBr.Visible = False
Call Me.dtgbrFF
Lb = UBound(Rb, 2) + 1
dtgBr.Rows = Lb + 50
dtgN.Rows = Lb + 50
For oo = 1 To Lb
    dtgBr.Row = oo
    dtgBr.Col = 1: dtgBr.Text = Rb(0, oo - 1)
    dtgBr.Col = 2: dtgBr.Text = Rb(1, oo - 1)
    dtgBr.Col = 3: dtgBr.Text = Rb(2, oo - 1)
    dtgBr.Col = 4: dtgBr.Text = Rb(3, oo - 1)
    
    dtgN.Row = oo
    dtgN.Col = 1: dtgN.Text = Rb(0, oo - 1)
    dtgN.Col = 2: dtgN.Text = Rb(1, oo - 1)
    dtgN.Col = 3: dtgN.Text = Rb(2, oo - 1)
    dtgN.Col = 4: dtgN.Text = Rb(3, oo - 1)
Next
dtgBr.Visible = True
End Sub

