VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmBxBrow 
   Caption         =   "���ı�����"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15210
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdFw 
      Caption         =   "ѡ����Ա"
      Height          =   315
      Left            =   8040
      TabIndex        =   27
      Top             =   8610
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "����"
      Height          =   585
      Left            =   14400
      Picture         =   "frmBxBrow.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8190
      Width           =   675
   End
   Begin VB.Frame Frame1 
      Caption         =   "���������뱨��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   8310
      TabIndex        =   25
      Top             =   480
      Width           =   7215
      Begin VB.CommandButton cmdNew 
         Caption         =   "��������"
         Height          =   705
         Left            =   5820
         TabIndex        =   26
         Top             =   750
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "��   ��"
      Height          =   315
      Left            =   8040
      TabIndex        =   24
      Top             =   0
      Width           =   2985
   End
   Begin VB.CommandButton cmdFF 
      Caption         =   "�·��ù���"
      Height          =   765
      Left            =   13170
      TabIndex        =   23
      Tag             =   "302"
      Top             =   2940
      Width           =   855
   End
   Begin VB.Frame frmYj 
      Caption         =   "��������"
      Height          =   3285
      Left            =   0
      TabIndex        =   13
      Top             =   5790
      Width           =   7995
      Begin VB.Frame Frame2 
         Caption         =   "������ѯ"
         Height          =   615
         Left            =   30
         TabIndex        =   17
         Top             =   2550
         Width           =   5745
         Begin VB.TextBox txtYc 
            Height          =   285
            Left            =   2820
            TabIndex        =   20
            Top             =   240
            Width           =   1635
         End
         Begin VB.ComboBox comXZ 
            Height          =   300
            ItemData        =   "frmBxBrow.frx":0102
            Left            =   810
            List            =   "frmBxBrow.frx":010F
            TabIndex        =   19
            Text            =   "��ͬ���"
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdRef1 
            Caption         =   "��  ѯ"
            Height          =   285
            Left            =   4590
            TabIndex        =   18
            Top             =   270
            Width           =   1035
         End
         Begin VB.Label Label2 
            Caption         =   "ֵ"
            Height          =   255
            Left            =   2610
            TabIndex        =   22
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label1 
            Caption         =   "����"
            Height          =   255
            Left            =   300
            TabIndex        =   21
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdYO 
         Caption         =   "��"
         Height          =   315
         Left            =   7110
         TabIndex        =   16
         Top             =   2820
         Width           =   795
      End
      Begin VB.CommandButton cmdBr 
         Caption         =   "ȫ����ʾ"
         Height          =   315
         Left            =   6120
         TabIndex        =   15
         Top             =   2820
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgYj 
         Height          =   2205
         Left            =   0
         TabIndex        =   14
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3889
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mga 
      Height          =   6615
      Left            =   -30
      TabIndex        =   4
      Top             =   -30
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   11668
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame frmOpt 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   435
      Left            =   450
      TabIndex        =   1
      Top             =   7470
      Width           =   7485
      Begin VB.OptionButton optQi 
         Caption         =   "����˹��ĵ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2700
         TabIndex        =   3
         Top             =   30
         Width           =   2025
      End
      Begin VB.OptionButton optMe 
         Caption         =   "�ҵı�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   300
         TabIndex        =   2
         Top             =   30
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame frmAdd 
      Height          =   6615
      Left            =   11010
      TabIndex        =   0
      Top             =   0
      Width           =   4245
      Begin VB.CommandButton cmdRight 
         Caption         =   "��һ��"
         Height          =   345
         Left            =   3330
         TabIndex        =   8
         Top             =   2370
         Width           =   825
      End
      Begin VB.CommandButton cmdLeft 
         Caption         =   "��һ��"
         Height          =   345
         Left            =   2430
         TabIndex        =   7
         Top             =   2370
         Width           =   855
      End
      Begin MSComCtl2.MonthView mtA 
         Height          =   2160
         Left            =   90
         TabIndex        =   6
         Top             =   120
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   3810
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   0
         MonthBackColor  =   -2147483633
         ShowToday       =   0   'False
         StartOfWeek     =   101056513
         TitleBackColor  =   16711935
         CurrentDate     =   38666
      End
      Begin VB.CommandButton cmdFyd 
         Height          =   795
         Index           =   0
         Left            =   330
         TabIndex        =   5
         Top             =   2970
         Width           =   915
      End
      Begin VB.Label lblFr 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2460
         Width           =   945
      End
      Begin VB.Label lblZZ 
         Caption         =   "~~"
         Height          =   165
         Left            =   1110
         TabIndex        =   10
         Top             =   2520
         Width           =   165
      End
      Begin VB.Label lblLr 
         Height          =   225
         Left            =   1350
         TabIndex        =   9
         Top             =   2460
         Width           =   945
      End
   End
   Begin VB.Label lblFw 
      Height          =   285
      Left            =   9240
      TabIndex        =   28
      Top             =   8610
      Width           =   1155
   End
End
Attribute VB_Name = "frmBxBrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AdoBxBro As ADODB.Recordset
Public adoYj As ADODB.Recordset

Private Sub cmdBack_Click()

frmBxBrow.Visible = False
frmZu.Enabled = True

frmZu.TBa.Buttons(3).Value = tbrUnpressed
End Sub

Private Sub cmdBr_Click()
Dim tt As String
On Error Resume Next
If mod1.KhK = 1 Then
    tt = "select * from newYjHt where bm='" & mod1.Bm & "' and ֧����=0  order by htrq desc"
ElseIf (mod1.KhK = 2 Or mod1.KhK = 3) And mod1.DName <> "�ܴ���" Then
    tt = "Select * from newYjht where comid=" & mod1.comId & " and ֧����=0  order by htrq desc"
ElseIf mod1.DName = "�ܴ���" Or mod1.DName = "�Ǽ���" Then
    tt = "Select * from newYjht where (֧����=0 or ֧����=1 and ��������>='" & DateSerial(Year(mod1.DQda), Month(mod1.DQda) - 1, 1) & "')   order by htrq desc"
End If
adoYj.Close
adoYj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgYJ.DataSource = adoYj

End Sub

Private Sub cmdFF_Click()
Dim Lb As Integer
On Error Resume Next

frmFYBX.Show

frmFYBX.frmYf.Visible = False
frmFYBX.cmdDao.Visible = False
frmFYBX.dtgNx.Visible = True
Call ModBx.FyQing

'Case 302
    Lb = 79
    frmFYBX.lblNlb.Caption = 79
    frmFYBX.cmdGui.Visible = False
'End Select

Call ModBx.dtgKj(Lb)

Dim tt As String
Dim ii As Integer
Dim TD As Date
Dim Tk As String
Dim TL As String

frmFYBX.Kd = True '���ο���,�Ա㱣��ʱ���ɿ�������
frmFYBX.LblTrq.Caption = mod1.DQda '��������
frmFYBX.lblBt.Caption = "�·��ù���"

frmFYBX.lblLcRen.Caption = mod1.DName
frmFYBX.lblLcUid.Caption = mod1.DHid

    '��������
    frmFYBX.comQy.Caption = mod1.Qy
    



 '��ҵ��Ա������
        
            
 
        frmFYBX.cmdAdd.Visible = True
        frmFYBX.cmdDel.Visible = True

      

                frmFYBX.cmdSave.Enabled = True


                frmFYBX.lblFr.Caption = lblFr.Caption
                frmFYBX.lblLr.Caption = lblLr.Caption
                frmFYBX.txtHg.Text = ""
                frmFYBX.lblDx.Caption = ""
                frmFYBX.lblBM.Caption = mod1.Bm
                frmFYBX.comQy.Caption = mod1.Qy
                frmFYBX.lblLc.Caption = 1 '���ο���,����Ϊ0
                frmFYBX.lblNewF.Caption = 1
               
                Set mod1.cmd = New ADODB.command
                mod1.cmd.ActiveConnection = mod1.cc
                mod1.cmd.CommandText = "FydAdd"
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@qy") = mod1.Qy
                mod1.cmd.Parameters("@bm") = mod1.Bm
                mod1.cmd.Parameters("@trq") = mod1.DQda
                mod1.cmd.Parameters("@ywy") = mod1.DName
                mod1.cmd.Parameters("@uid") = mod1.DHid
                mod1.cmd.Parameters("@Lcou") = 4 '��������
                mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
                mod1.cmd.Parameters("@lcRen") = mod1.DName
                mod1.cmd.Parameters("@lcUid") = mod1.DHid
                mod1.cmd.Parameters("@Lb") = Lb
                mod1.cmd.Execute
                frmFYBX.lblBh.Caption = mod1.cmd.Parameters("@bxid").Value
                Set cmd = Nothing
               
               
                tt = "fydAddB(" & frmFYBX.lblBh.Caption & ")"
                frmFYBX.adoF2.Recordset.Close
                frmFYBX.adoF2.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
                
'                frmFYBX.adoF2.Recordset.AddNew "BM", frmZu.lblBM.Caption
'                frmFYBX.adoF2.Recordset.Update "qy", comQy.Text
'                frmFYBX.adoF2.Recordset.Update "ywy", frmZu.comRen.Text
'                frmFYBX.adoF2.Recordset.Update "bxId", frmFYBX.lblBh.Caption
'                frmFYBX.adoF2.Recordset.Update "XG", 0
                Set frmFYBX.dtgBx.DataSource = frmFYBX.adoF2
                frmFYBX.dtgBx.AllowUpdate = True
                frmFYBX.txtBz.Enabled = True
                frmFYBX.txtBz.Locked = False
                
                frmFYBX.cmdAdd.Enabled = True
                frmFYBX.cmdDel.Enabled = True

                
        'End If


frmFYBX.lblYwy.Caption = mod1.DName
frmFYBX.lblUid.Caption = mod1.DHid
frmFYBX.frmRen.Visible = True

        tt = "Select atime as ����,khmc as ��������,sj as ����,fwbt as ���ݲ���,lyf as ���η�,gwf as ���·�,txf as ͨ�ŷ�,njtf as ���ڽ�ͨ��,wjtf as ���⽻ͨ��," & _
        "tcf as ͣ����,clf as ������,yf as �˷�,zcf as ס�޷�,bmtd as �����Ŷӷ�,cf as �ͷ�,ZDF as �д���,LPF as ��Ʒ��,fz as ����,WYF as ��ҵ��," & _
        "sd as ˮ��,DW as �绰,BGYP as �칫��Ʒ,YZ as ����,SZTG as �г��ƹ�,RYZP as ��Ա��Ƹ,KDF as ��ݷ�,PXF as ��ѵ��,CWSX as ����������,TDJS as �Ŷӽ����," & _
        "GTCF as ����ͣ����,GCLF as ����������,gg as ����,yH as �׺�,wl as ����,qtF as ������,gjj as ������,zhbx as �ۺϱ���,jtbt as ��ͨ����,zwbt as פ�����,gwbt as ��λ����,bm as ����,qy as ����,ywy as ����," & _
        "bid,gzdh as ���⳵ע�� from fyBx where Bxid=" & Val(frmFYBX.lblBh.Caption) & " order by bm,bid"
        frmFYBX.Fmx.Close
        frmFYBX.Fmx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Call ModBx.DiZ

'�������̰�ť
Call ModBx.AddLcBut(Lb)

frmFYBX.cmdGui.Visible = True

    frmFYBX.dtgBx.Columns("���ݲ���").Visible = False
    frmFYBX.txtQc.Enabled = False
    frmFYBX.frmQm.Enabled = True
    frmFYBX.frmNewQ.Visible = True
    frmFYBX.ZOrder 0
    frmFYBX.frmED.Visible = True
    frmFYBX.cmdAdd.Visible = False
    frmFYBX.cmdDel.Visible = False
    frmFYBX.lblGZDH.Visible = False
    frmFYBX.txtGZDH.Visible = False
    frmFYBX.opt1.Value = False
    frmFYBX.opt2.Value = False
'    If mod1.BM = "���̲�" Then
'        frmFYBX.lblGZDH.Visible = True
'        frmFYBX.txtGZDH.Visible = True
'    End If
End Sub

Private Sub cmdFw_Click()
Set Ren.XForm = New frmBxBrow
Call mod1.RenXz("frmBxBrow", Me, 0)
End Sub

Private Sub cmdFyd_Click(Index As Integer)
Dim Lb As Integer
On Error Resume Next

'���ж��Ƿ������Ϊ������С��λ�����˱���
Dim TG As Integer
TG = cmdFyd(Index).Tag
If TG = 29 Or TG = 33 Or TG = 51 Or TG = 54 Or TG = 64 Or TG = 197 Or TG = 206 Then
    If mod1.FYF = False Then
        MsgBox "���������Ƿ�����С��λ,���Բ���������͵�,������������ϵ!"
        Exit Sub
    End If
End If

frmFYBX.Show
frmFYBX.frmED.Visible = False
frmFYBX.dtgNx.Visible = False
frmFYBX.cmdGui.Visible = False
frmFYBX.frmYf.Visible = False
frmFYBX.cmdDao.Visible = False
Call ModBx.FyQing

Select Case cmdFyd(Index).Tag
Case 29 '��������
    Lb = 7
    frmFYBX.lblNlb.Caption = 7
Case 33 '�ܾ�����
    Lb = 8
    frmFYBX.lblNlb.Caption = 8
Case 183  '�˷�
    Lb = 50
    frmFYBX.lblNlb.Caption = 50
    frmFYBX.frmYf.Visible = True
Case 148 '����
    Lb = 35
    frmFYBX.lblNlb.Caption = 35
    frmFYBX.cmdGui.Visible = True
    frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
    frmFYBX.lblNewF = 1
Case 42 '�������(>500)
    Lb = 11
    frmFYBX.frmWd.Visible = True
    frmFYBX.comYwy.Enabled = False
    frmFYBX.comXmmc.Enabled = False
    frmFYBX.lblNlb.Caption = 11
Case 51 '���۾���
    Lb = 13
    frmFYBX.lblNlb.Caption = 13
Case 54 '���ž���
    Lb = 14
    frmFYBX.lblNlb.Caption = 14
Case 57 'ҵ��Ա(>1000)
    Lb = 15
    frmFYBX.lblNlb.Caption = 15
    frmFYBX.dtgNx.Visible = False
    frmFYBX.lblNewF = 1
    'ҵ��Ա�Զ����ɱ���.
Case 61
    Lb = 16
    frmFYBX.lblNlb.Caption = 16
    frmFYBX.dtgNx.Visible = False
Case 64 '��ͨ����(>1000)
    Lb = 17
    frmFYBX.lblNlb.Caption = 17
Case 77 '�����Ŷ�(>500)
    Lb = 20
    frmFYBX.lblNlb.Caption = 20
Case 136 '���ù���
    Lb = 32
    frmFYBX.lblNlb.Caption = 32
    frmFYBX.cmdGui.Visible = True
Case 197 '���۾���
    Lb = 53
    frmFYBX.lblNlb.Caption = 53
Case 206 '���̲�
    Lb = 54
    frmFYBX.lblNlb.Caption = 54
Case 211 '����
    Lb = 55
    frmFYBX.lblNlb.Caption = 55
    frmFYBX.cmdGui.Visible = True
    frmFYBX.cmdDao.Visible = True
    frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
    frmFYBX.lblNewF = 1
Case 215 '������
    Lb = 56
    frmFYBX.lblNlb.Caption = 56
    frmFYBX.cmdGui.Visible = True
    frmFYBX.cmdDao.Visible = True
    frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
    frmFYBX.lblNewF = 1
Case 223 '���´���������
    Lb = 58
    frmFYBX.lblNlb.Caption = 58
Case 227                '������Ա�ۺϱ���
    Lb = 59
    frmFYBX.lblNlb.Caption = 59
    frmFYBX.cmdDao.Visible = True
    frmFYBX.cmdGui.Visible = True
    frmFYBX.lblNewF = 1
Case 285
    Lb = 72
    frmFYBX.lblNlb.Caption = 72
Case 314
    Lb = 82
    frmFYBX.lblNlb.Caption = 82
Case 322
    Lb = 84
    frmFYBX.lblNlb.Caption = 84
    'frmFYBX.cmdG.Visible = True
    frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
    frmFYBX.lblNewF = 1
End Select

Call ModBx.dtgKj(Lb)

Dim tt As String
Dim ii As Integer
Dim TD As Date
Dim Tk As String
Dim TL As String

frmFYBX.Kd = True '���ο���,�Ա㱣��ʱ���ɿ�������
frmFYBX.LblTrq.Caption = mod1.DQda '��������
frmFYBX.lblBt.Caption = cmdFyd(Index).Caption

frmFYBX.lblLcRen.Caption = mod1.DName
frmFYBX.lblLcUid.Caption = mod1.DHid

    '��������
    frmFYBX.comQy.Caption = mod1.Qy
    
If cmdFyd(Index).Tag = 57 Or cmdFyd(Index).Tag = 61 Then           'ҵ��Ա�Զ����ɱ���.
    frmFYBX.lblNewF.Caption = 1
    frmFYBX.lblGui.Caption = mod1.DName
    frmFYBX.lblGuid.Caption = mod1.DHid
        '���ж�FyD�����Ƿ��д˼�¼�����������������û�У������FyTG�������µ���
    Call mod1.WeeKDay(mtA.Value)

    tt = "Select nlb,bxid from Fyd where ywy='" & mod1.DName & "' and fRq='" & mod1.FR & "' and lRq='" & mod1.LR & "' and Nlb=" & Lb
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    
    If mod1.HTP.RecordCount > 0 Then
'        Call modXmGz.fyBound
'        frmFYBX.cmdSave.Enabled = False
        frmFYBX.Visible = False
        MsgBox ("���ڱ������Ѿ�����,���Ҫ�޸�,����ɾ���˱�����!")
        frmBxBrow.Enabled = True
        Exit Sub
    Else '�����µı�����
    
        '�����ܱ�
'        If mod1.KhK = 1 Then '�Ƿ�Ϊ���۾���
'            frmFYBX.lblFxz = 6
'        Else
'            frmFYBX.lblFxz = 0
'        End If
        frmFYBX.lblFr.Caption = lblFr.Caption
        frmFYBX.lblLr.Caption = lblLr.Caption
        frmFYBX.lblBM.Caption = mod1.Bm
        frmFYBX.comQy.Caption = mod1.Qy
        frmFYBX.lblLc.Caption = 0 '���ο���,����Ϊ0
        Set mod1.cmd = New ADODB.command
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "FydAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@qy") = mod1.Qy
        mod1.cmd.Parameters("@bm") = mod1.Bm
        mod1.cmd.Parameters("@trq") = mod1.DQda
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@Lcou") = Right(cmdFyd(Index).ToolTipText, 1) '��������
        mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
        mod1.cmd.Parameters("@lcRen") = mod1.DName
        mod1.cmd.Parameters("@lcUid") = mod1.DHid
        mod1.cmd.Parameters("@Lb") = Lb
        mod1.cmd.Execute
        frmFYBX.lblBh.Caption = mod1.cmd.Parameters("@bxid").Value
        Set cmd = Nothing
                
        tt = "Select * from fyBx where Bid=99999"
        
        frmFYBX.adoF2.Recordset.Close
        frmFYBX.adoF2.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        Set frmFYBX.dtgBx.DataSource = frmFYBX.adoF2
        

        frmFYBX.cmdSave.Enabled = True

        
        frmFYBX.lblFr.Caption = mod1.FR
        frmFYBX.lblLr.Caption = mod1.LR
        frmFYBX.txtHg.Text = ""
        frmFYBX.lblDx.Caption = ""
        
        '������ϸ��
'        tt = "Select * from fyTgP where ywy ='" & mod1.DName & "' and aTime>='" & mod1.Fr & "' and aTime<='" & mod1.lr & _
        "' order by aTime,khmc"
        '������ϸ��
        tt = "Select * from fyTg where ywy ='" & mod1.DName & "' and aTime>='" & mod1.FR & "' and aTime<'" & mod1.LR & _
        "' order by aTime,khmc"
        frmFYBX.adoFy.Recordset.Close
        frmFYBX.adoFy.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        'Set frmFYBX.dtgFy.DataSource = frmFYBX.adoFy
           

        
        
           '������ϸ������ܱ�
        frmFYBX.adoFy.Recordset.MoveFirst
        TD = frmFYBX.adoFy.Recordset.Fields("aTime").Value
        Tk = frmFYBX.adoFy.Recordset.Fields("khmc").Value
        frmFYBX.adoF2.Recordset.AddNew "aTime", frmFYBX.adoFy.Recordset.Fields("aTime").Value
        frmFYBX.adoF2.Recordset.Update "Fid", frmFYBX.adoFy.Recordset.Fields("Fid").Value
        frmFYBX.adoF2.Recordset.Update "bxid", Val(frmFYBX.lblBh.Caption)
        Do While Not frmFYBX.adoFy.Recordset.EOF
            
                'ת����Ŀ���ֵ
                Select Case frmFYBX.adoFy.Recordset.Fields("fylb").Value
                Case "���ڽ�ͨ��"
                TL = "NJTF"
                Case "���⽻ͨ��"
                TL = "WJTF"
                Case "ס�޷�"
                TL = "ZCF"
                Case "�ͷ�"
                TL = "CF"
                Case "�д���"
                TL = "ZDF"
                Case "��Ʒ��"
                TL = "LPF"
                Case "��ݷ�"
                TL = "KDF"
                Case "ͨ�ŷ�"
                TL = "TXF"
                Case "������"
                TL = "CLF"
        '        Case "С��"
        '        TL = "xg"
                End Select
                If frmFYBX.adoFy.Recordset.Fields("fy").Value > 1000 And Lb = 15 Then
                    If TD = frmFYBX.adoFy.Recordset.Fields("aTime").Value And Tk = frmFYBX.adoFy.Recordset.Fields("khmc").Value Then
                        frmFYBX.adoF2.Recordset.Update "aTime", frmFYBX.adoFy.Recordset.Fields("aTime").Value
                        frmFYBX.adoF2.Recordset.Update "khmc", frmFYBX.adoFy.Recordset.Fields("khmc").Value
                        frmFYBX.adoF2.Recordset.Update "ywy", mod1.DName
                        frmFYBX.adoF2.Recordset.Update "ywyuid", mod1.DHid
                        frmFYBX.adoF2.Recordset.Update "bm", mod1.Bm
                        frmFYBX.adoF2.Recordset.Update "qy", mod1.Qy
                        If IsNull(frmFYBX.adoF2.Recordset.Fields(TL).Value) = True Then frmFYBX.adoF2.Recordset.Fields(TL).Value = 0
                        frmFYBX.adoF2.Recordset.Update TL, _
                        (frmFYBX.adoF2.Recordset.Fields(TL).Value + frmFYBX.adoFy.Recordset.Fields("fy").Value)
                        frmFYBX.adoF2.Recordset.Update "bxid", Val(frmFYBX.lblBh.Caption)
                        frmFYBX.adoF2.Recordset.Update "GongF", 2
                    Else
                        If Not (frmFYBX.adoF2.Recordset.Fields("xg").Value > 0) Then
                            frmFYBX.adoF2.Recordset.Delete adAffectCurrent
                        End If
                        frmFYBX.adoF2.Recordset.AddNew "aTime", frmFYBX.adoFy.Recordset.Fields("aTime").Value
                        frmFYBX.adoF2.Recordset.Update "khmc", frmFYBX.adoFy.Recordset.Fields("khmc").Value
                        frmFYBX.adoF2.Recordset.Update "bm", mod1.Bm
                        frmFYBX.adoF2.Recordset.Update "qy", mod1.Qy
                        frmFYBX.adoF2.Recordset.Update "ywy", mod1.DName
                        frmFYBX.adoF2.Recordset.Update "ywyuid", mod1.DHid
                        frmFYBX.adoF2.Recordset.Update TL, frmFYBX.adoFy.Recordset.Fields("fy").Value
                        frmFYBX.adoF2.Recordset.Update "bxid", Val(frmFYBX.lblBh.Caption)
                        frmFYBX.adoF2.Recordset.Update "GongF", 2
                        TD = frmFYBX.adoFy.Recordset.Fields("aTime").Value
                        Tk = frmFYBX.adoFy.Recordset.Fields("khmc").Value
                    End If
                    '����С��
                    If IsNull(frmFYBX.adoF2.Recordset.Fields("xg").Value) = True Then
                        frmFYBX.adoF2.Recordset.Fields("xg").Value = 0
                    End If
                    frmFYBX.adoF2.Recordset.Fields("xg").Value = frmFYBX.adoF2.Recordset.Fields("xg").Value + frmFYBX.adoFy.Recordset.Fields("fy").Value
                    frmFYBX.txtHg.Text = Val(frmFYBX.txtHg.Text) + frmFYBX.adoFy.Recordset.Fields("fy").Value
                ElseIf frmFYBX.adoFy.Recordset.Fields("fy").Value <= 1000 And Lb = 16 Then
                    If TD = frmFYBX.adoFy.Recordset.Fields("aTime").Value And Tk = frmFYBX.adoFy.Recordset.Fields("khmc").Value Then
                        frmFYBX.adoF2.Recordset.Update "aTime", frmFYBX.adoFy.Recordset.Fields("aTime").Value
                        frmFYBX.adoF2.Recordset.Update "khmc", frmFYBX.adoFy.Recordset.Fields("khmc").Value
                        frmFYBX.adoF2.Recordset.Update "bm", mod1.Bm
                        frmFYBX.adoF2.Recordset.Update "qy", mod1.Qy
                        frmFYBX.adoF2.Recordset.Update "ywy", mod1.DName
                        frmFYBX.adoF2.Recordset.Update "ywyuid", mod1.DHid
                        If IsNull(frmFYBX.adoF2.Recordset.Fields(TL).Value) = True Then frmFYBX.adoF2.Recordset.Fields(TL).Value = 0
                        frmFYBX.adoF2.Recordset.Update TL, _
                        (frmFYBX.adoF2.Recordset.Fields(TL).Value + frmFYBX.adoFy.Recordset.Fields("fy").Value)
                        frmFYBX.adoF2.Recordset.Update "bxid", Val(frmFYBX.lblBh.Caption)
                        frmFYBX.adoF2.Recordset.Update "GongF", 2
                    Else
                 
                        frmFYBX.adoF2.Recordset.AddNew "aTime", frmFYBX.adoFy.Recordset.Fields("aTime").Value
                        frmFYBX.adoF2.Recordset.Update "khmc", frmFYBX.adoFy.Recordset.Fields("khmc").Value
                        frmFYBX.adoF2.Recordset.Update "bm", mod1.Bm
                        frmFYBX.adoF2.Recordset.Update "qy", mod1.Qy
                        frmFYBX.adoF2.Recordset.Update "ywy", mod1.DName
                        frmFYBX.adoF2.Recordset.Update "ywyuid", mod1.DHid
                        frmFYBX.adoF2.Recordset.Update TL, frmFYBX.adoFy.Recordset.Fields("fy").Value
                        frmFYBX.adoF2.Recordset.Update "bxid", Val(frmFYBX.lblBh.Caption)
                        frmFYBX.adoF2.Recordset.Update "GongF", 2
                        TD = frmFYBX.adoFy.Recordset.Fields("aTime").Value
                        Tk = frmFYBX.adoFy.Recordset.Fields("khmc").Value
                    End If
                    '����С��
                    If IsNull(frmFYBX.adoF2.Recordset.Fields("xg").Value) = True Then
                        frmFYBX.adoF2.Recordset.Fields("xg").Value = 0
                    End If
                    frmFYBX.adoF2.Recordset.Fields("xg").Value = frmFYBX.adoF2.Recordset.Fields("xg").Value + frmFYBX.adoFy.Recordset.Fields("fy").Value
                    frmFYBX.txtHg.Text = Val(frmFYBX.txtHg.Text) + frmFYBX.adoFy.Recordset.Fields("fy").Value
                End If
            frmFYBX.adoFy.Recordset.MoveNext
        Loop
            frmFYBX.lblDx.Caption = mod1.ChangBi(frmFYBX.txtHg.Text)

            Set frmFYBX.dtgBx.DataSource = frmFYBX.adoF2
          
   
    End If
            frmFYBX.dtgBx.Visible = True
            frmFYBX.dtgNx.Visible = False
            frmFYBX.lblBM.Caption = mod1.Bm
            frmFYBX.lblGui.Caption = mod1.DName
            frmFYBX.lblGuid.Caption = mod1.DHid
            Set frmFYBX.dtgBx.DataSource = frmFYBX.adoF2
            

        
            frmFYBX.cmdAdd.Enabled = False
            frmFYBX.cmdDel.Enabled = False
            frmFYBX.dtgBx.AllowUpdate = True
            frmFYBX.txtBz.Locked = False
            frmFYBX.txtBz.Enabled = True

'''        tt = "Select atime as ����,khmc as ��������,sj as �Ľ�,fwbt as ���ݲ���,lyf as ���η�,gwf as ���·�,txf as ͨ�ŷ�,njtf as ���ڽ�ͨ��,wjtf as ���⽻ͨ��," & _
'''        "tcf as ͣ����,clf as ������,yf as �˷�,zcf as ס�޷�,bmtd as �����Ŷӷ�,cf as �ͷ�,ZDF as �д���,LPF as ��Ʒ��,fz as ����,WYF as ��ҵ��," & _
'''        "sd as ˮ��,DW as �绰,BGYP as �칫��Ʒ,YZ as ����,SZTG as �г��ƹ�,RYZP as ��Ա��Ƹ,KDF as ��ݷ�,PXF as ��ѵ��,CWSX as ����������,TDJS as �Ŷӽ����," & _
'''        "qtf as ����,GTCF as ����ͣ����,GCLF as ����������,gg as ����,yH as �׺�,wl as ����,sj as ������,bm as ����,qy as ����,ywy as ����," & _
'''        "bid,gzdh as ���⳵ע�� from fyBx where Bxid=" & Val(frmFYBX.lblBh.Caption) & " order by bm,bid"
'''        frmFYBX.Fmx.Close
'''        frmFYBX.Fmx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''        Call ModBx.DiZ
'''            frmFYBX.dtgBx.Visible = False
'''            frmFYBX.dtgNx.Visible = True
            MsgBox "��ע���˱������еĳ��⳵���ã�"


Else '��ҵ��Ա������
        
            
 
        frmFYBX.cmdAdd.Visible = True
        frmFYBX.cmdDel.Visible = True

      

                frmFYBX.cmdSave.Enabled = True


                frmFYBX.lblFr.Caption = lblFr.Caption
                frmFYBX.lblLr.Caption = lblLr.Caption
                frmFYBX.txtHg.Text = ""
                frmFYBX.lblDx.Caption = ""
                frmFYBX.lblBM.Caption = mod1.Bm
                frmFYBX.comQy.Caption = mod1.Qy
                frmFYBX.lblLc.Caption = 0 '���ο���,����Ϊ0
               
                Set mod1.cmd = New ADODB.command
                mod1.cmd.ActiveConnection = mod1.cc
                mod1.cmd.CommandText = "FydAdd"
                mod1.cmd.CommandType = adCmdStoredProc
                mod1.cmd.Parameters("@qy") = mod1.Qy
                mod1.cmd.Parameters("@bm") = mod1.Bm
                mod1.cmd.Parameters("@trq") = mod1.DQda
                mod1.cmd.Parameters("@ywy") = mod1.DName
                mod1.cmd.Parameters("@uid") = mod1.DHid
                mod1.cmd.Parameters("@Lcou") = Right(cmdFyd(Index).ToolTipText, 1) '��������
                mod1.cmd.Parameters("@Lc") = 0 '��ǰ����
                mod1.cmd.Parameters("@lcRen") = mod1.DName
                mod1.cmd.Parameters("@lcUid") = mod1.DHid
                mod1.cmd.Parameters("@Lb") = Lb
                mod1.cmd.Execute
                frmFYBX.lblBh.Caption = mod1.cmd.Parameters("@bxid").Value
                Set cmd = Nothing
               
               
                tt = "fydAddB(" & frmFYBX.lblBh.Caption & ")"
                frmFYBX.adoF2.Recordset.Close
                frmFYBX.adoF2.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
                
'                frmFYBX.adoF2.Recordset.AddNew "BM", frmZu.lblBM.Caption
'                frmFYBX.adoF2.Recordset.Update "qy", comQy.Text
'                frmFYBX.adoF2.Recordset.Update "ywy", frmZu.comRen.Text
'                frmFYBX.adoF2.Recordset.Update "bxId", frmFYBX.lblBh.Caption
'                frmFYBX.adoF2.Recordset.Update "XG", 0
                Set frmFYBX.dtgBx.DataSource = frmFYBX.adoF2
                
                frmFYBX.dtgBx.AllowUpdate = True
                frmFYBX.txtBz.Enabled = True
                frmFYBX.txtBz.Locked = False
                
                frmFYBX.cmdAdd.Enabled = True
                frmFYBX.cmdDel.Enabled = True

                
        'End If


End If
frmFYBX.lblYwy.Caption = mod1.DName
frmFYBX.lblUid.Caption = mod1.DHid


        tt = "Select atime as ����,khmc as ��������,sj as ����,fwbt as ���ݲ���,lyf as ���η�,gwf as ���·�,txf as ͨ�ŷ�,njtf as ���ڽ�ͨ��,wjtf as ���⽻ͨ��," & _
        "tcf as ͣ����,clf as ������,yf as �˷�,zcf as ס�޷�,bmtd as �����Ŷӷ�,cf as �ͷ�,ZDF as �д���,LPF as ��Ʒ��,fz as ����,WYF as ��ҵ��," & _
        "sd as ˮ��,DW as �绰,BGYP as �칫��Ʒ,YZ as ����,SZTG as �г��ƹ�,RYZP as ��Ա��Ƹ,KDF as ��ݷ�,PXF as ��ѵ��,CWSX as ����������,TDJS as �Ŷӽ����," & _
        "GTCF as ����ͣ����,GCLF as ����������,gg as ����,yH as �׺�,wl as ����,qtF as ������,gjj as ������,zhbx as �ۺϱ���,jtbt as ��ͨ����,zwbt as פ�����,gwbt as ��λ����,bm as ����,qy as ����,ywy as ����," & _
        "bid,gzdh as ���⳵ע�� from fyBx where Bxid=" & Val(frmFYBX.lblBh.Caption) & " order by bm,bid"
        frmFYBX.Fmx.Close
        frmFYBX.Fmx.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Call ModBx.DiZ
If frmFYBX.lblNlb.Caption = 35 Then '�����ѳ��ο��������Ա༭����
    frmFYBX.dtgBx.Visible = True
    frmFYBX.dtgNx.Visible = False
    frmFYBX.cmdDao.Visible = True
End If

'�������̰�ť
Call ModBx.AddLcBut(Lb)



    'frmFYBX.dtgBx.Columns("���ݲ���").Visible = False
    frmFYBX.txtQc.Enabled = False
    frmFYBX.frmQm.Enabled = True
    frmFYBX.frmNewQ.Visible = True
    frmFYBX.ZOrder 0
End Sub

Private Sub cmdLeft_Click()
Dim Ldate As Date
On Error Resume Next
Ldate = mtA.Value
mtA.Value = DateSerial(Year(Ldate), Month(Ldate), Day(Ldate) - 7)
Call GetWeek

End Sub

Private Sub cmdNew_Click()
FYDN1.Show
FYDN1.ZOrder 0
End Sub

Private Sub cmdOpen_Click()
mga.Col = 4
'MsgBox MGa.Text

If Val(mga.Text) = 0 Then Exit Sub
If mod1.DKZ(mga.Text, 2) = True Then
        MsgBox "��ݱ�����" & mod1.DKRen & "��,���Ժ�����,������������ϵ."
        Exit Sub
End If

frmBxBrow.Enabled = False
frmFYBX.Show

Call ModBx.FyQing
Call ModBx.fydBound(Val(mga.Text))

End Sub

Private Sub cmdRef1_Click()
Dim tt As String
On Error Resume Next
Select Case comXZ.Text
    Case "��ͬ���"
        If mod1.KhK = 1 And mod1.BmJl = True And mod1.DName <> "���Ƿ�" Then
            tt = "select * from newYjHt where (ggl='" & mod1.DHid & "' or bm='" & mod1.Bm & "' or lcren='" & mod1.DName & "') and ��ͬ���=" & Val(txtYc.Text) & " and (֧����=0 or ֧���� is null) "
'''''        ElseIf (mod1.KhK = 2 Or mod1.KhK = 3) And mod1.DName <> "�ܴ���" Then
'''''            tt = "select * from newYjHt where ��ͬ���=" & Val(txtYc.Text) & " and (֧����=0 or ֧���� is null) order by htrq desc"
        ElseIf mod1.DName = "�Ǽ���" Or mod1.DName = "������" Then
            'tt = "select * from newYjHtZ where ��ͬ���=" & Val(txtYc.Text) & " and (֧����=0 or ֧���� is null or ֧����=1 and ��������>='" & DateSerial(Year(mod1.DQda), Month(mod1.DQda) - 1, 1) & "')  order by htrq desc"
            tt = "select * from newYjHtZ where ��ͬ���=" & Val(txtYc.Text) & " and (֧����=0 or ֧���� is null or ֧����=1 )  order by htrq desc"
        ElseIf mod1.DName = "���Ƿ�" Then
            tt = "select * from newYjHtZ where ��ͬ���=" & Val(txtYc.Text) & " and (֧����=0 or ֧���� is null or ֧����=1 ) and (qy='�Ϻ�' or qy='����' or qy='�Ͼ�' or qy='��̨') order by htrq desc"
        End If
        
    Case "��Ŀ����"
        If mod1.KhK = 1 And mod1.BmJl = True And mod1.DName <> "���Ƿ�" Then
            tt = "select * from newYjHt where (ggl='" & mod1.DHid & "' or bm='" & mod1.Bm & "'  or lcren='" & mod1.DName & "') and ��Ŀ���� like '%" & Trim(txtYc.Text) & "%'" & " and (֧����=0 or ֧���� is null)"
'''''        ElseIf (mod1.KhK = 2 Or mod1.KhK = 3) And mod1.DName <> "�ܴ���" Then
'''''            tt = "select * from newYjHt where ��Ŀ���� like '%" & Trim(txtYc.Text) & "%'" & " and (֧����=0 or ֧���� is null) order by htrq desc"
        ElseIf mod1.DName = "�Ǽ���" Or mod1.DName = "������" Then
            'tt = "select * from newYjHtZ where ��Ŀ���� like '%" & Trim(txtYc.Text) & "%' and (֧����=0 or ֧���� is null or ֧����=1 and ��������>='" & DateSerial(Year(mod1.DQda), Month(mod1.DQda) - 1, 1) & "')   order by htrq desc"
            tt = "select * from newYjHtZ where ��Ŀ���� like '%" & Trim(txtYc.Text) & "%' and (֧����=0 or ֧���� is null or ֧����=1)   order by htrq desc"
        ElseIf mod1.DName = "���Ƿ�" Then
            tt = "select * from newYjHtZ where ��Ŀ���� like '%" & Trim(txtYc.Text) & "%' and (֧����=0 or ֧���� is null or ֧����=1) and (qy='�Ϻ�' or qy='����' or qy='�Ͼ�' or qy='��̨')    order by htrq desc"
        End If
    Case "��ͬ���"
        If mod1.KhK = 1 And mod1.BmJl = True And mod1.DName <> "���Ƿ�" Then
            tt = "select * from newYjHt where ��ͬ��� like '%" & Trim(txtYc.Text) & "%' and (֧����=0 or ֧���� is null) and  (ggl='" & mod1.DHid & "' or bm='" & mod1.Bm & "' or lcren='" & mod1.DName & "')"
            'tt = "select * from newYjHtZ where ��ͬ��� like '%" & Trim(txtYc.Text) & "%' and (֧����=0 or ֧���� is null or ֧����=1 )   order by htrq desc"
'''''        ElseIf (mod1.KhK = 2 Or mod1.KhK = 3) And mod1.DName <> "�ܴ���" Then
'''''            tt = "select * from newYjHt where ��ͬ��� like '%" & Trim(txtYc.Text) & "%'" & " and (֧����=0 or ֧���� is null) order by htrq desc"
        ElseIf mod1.DName = "�Ǽ���" Or mod1.DName = "������" Then
            'tt = "select * from newYjHtZ where ��ͬ��� like '%" & Trim(txtYc.Text) & "%' and (֧����=0 or ֧���� is null or ֧����=1 and ��������>='" & DateSerial(Year(mod1.DQda), Month(mod1.DQda) - 1, 1) & "')   order by htrq desc"
            tt = "select * from newYjHtZ where ��ͬ��� like '%" & Trim(txtYc.Text) & "%' and (֧����=0 or ֧���� is null or ֧����=1 )   order by htrq desc"
        ElseIf mod1.DName = "���Ƿ�" Then
            tt = "select * from newYjHtZ where ��ͬ��� like '%" & Trim(txtYc.Text) & "%' and (֧����=0 or ֧���� is null or ֧����=1 ) and (qy='�Ϻ�' or qy='����' or qy='�Ͼ�' or qy='��̨')       order by htrq desc"
        End If
End Select
adoYj.Close
adoYj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgYJ.DataSource = adoYj
If adoYj.RecordCount = 0 Then
    Set dtgYJ.DataSource = adoYj
    dtgYJ.Rows = 2
    dtgYJ.FixedRows = 0
    dtgYJ.FixedRows = 1

Else
    dtgYJ.Rows = 2
    dtgYJ.FixedRows = 1
    Set dtgYJ.DataSource = adoYj
End If

dtgYJ.Row = dtgYJ.Rows
End Sub

Private Sub cmdRight_Click()
Dim Ldate As Date
On Error Resume Next
Ldate = mtA.Value
mtA.Value = DateSerial(Year(Ldate), Month(Ldate), Day(Ldate) + 7)
Call GetWeek
End Sub


Private Sub cmdYO_Click()
Dim tt As String
Dim oo As Integer
Dim Pwf As Boolean
Dim QFF As Boolean '��ͬȫ��֧����
Dim Ny As Single '��֧���Ľ����ܶ�(�°��е�,����÷�������е�)

Dim Yid As Long
Dim Xmmc As String
Dim Htbh As String
dtgYJ.Col = 11
Yid = Val(dtgYJ.Text)
dtgYJ.Col = 3
Xmmc = dtgYJ.Text
dtgYJ.Col = 5
Htbh = dtgYJ.Text
'������֤
On Error GoTo YZERR
If Htbh = "" Then
    Exit Sub
End If
tt = "insert into HMText.dbo.ML (NB,NBLX,trq,bh,ywy,uid,Bz,mt3) values ('����','�鿴',getdate(),'" & Yid & "','" & mod1.DName & "','" & mod1.DHid & "' ,'" & Xmmc & "','" & Htbh & "')"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
'mod1.HTP.Close
Set mod1.HTP = Nothing


On Error Resume Next
'''''''''Call frmYjBx.yjBXQing
'''''''''QFF = False
'''''''''dtgYj.Col = 1
'''''''''frmYjBx.lblQy.Caption = dtgYj.Text
'''''''''dtgYj.Col = 2
'''''''''frmYjBx.lblBm.Caption = dtgYj.Text
'''''''''dtgYj.Col = 3
'''''''''frmYjBx.lblXmmc.Caption = dtgYj.Text
'''''''''dtgYj.Col = 5
'''''''''frmYjBx.lblHtbh.Text = dtgYj.Text
'''''''''dtgYj.Col = 4
'''''''''frmYjBx.lblHtZe.Caption = dtgYj.Text
'''''''''dtgYj.Col = 7
'''''''''frmYjBx.lblYf.Caption = dtgYj.Text
'''''''''dtgYj.Col = 6
'''''''''frmYjBx.lblED.Caption = dtgYj.Text
'''''''''dtgYj.Col = 23
'''''''''QFF = dtgYj.Text
'''''''''If QFF = True Then
'''''''''    frmYjBx.lblQFF.Caption = "ȫ��֧�����"
'''''''''    frmYjBx.lblQFF.ForeColor = &HFF&
'''''''''Else
'''''''''    frmYjBx.lblQFF.Caption = "δ���"
'''''''''End If
'''''''''dtgYj.Col = 11
'''''''''frmYjBx.lblYid.Caption = dtgYj.Text
'''''''''dtgYj.Col = 14
'''''''''frmYjBx.lblYwy.Caption = dtgYj.Text
'''''''''dtgYj.Col = 15
'''''''''frmYjBx.lblUid.Caption = dtgYj.Text
'''''''''dtgYj.Col = 16
'''''''''frmYjBx.lblLc.Caption = dtgYj.Text
'''''''''dtgYj.Col = 17
'''''''''frmYjBx.lblLcRen.Caption = dtgYj.Text
'''''''''dtgYj.Col = 18
'''''''''frmYjBx.lblLcUid.Caption = dtgYj.Text
'''''''''dtgYj.Col = 19
'''''''''frmYjBx.lblFwid.Caption = dtgYj.Text
'''''''''dtgYj.Col = 21
'''''''''frmYjBx.txtCXF.Text = dtgYj.Text
'''''''''dtgYj.Col = 22
'''''''''Pwf = dtgYj.Text
'''''''''dtgYj.Col = 10
'''''''''frmYjBx.txtBz.Text = dtgYj.Text
'''''''''Ny = 0
'''''''''tt = "select yj from htping where htbh='" & frmYjBx.lblHtbh.Text & "'"
'''''''''Set mod1.HTP = New ADODB.Recordset
'''''''''mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''frmYjBx.lblYj.Caption = mod1.HTP.Fields("yj").Value
'''''''''tt = "select sum(Ӧ��)+sum(cxf) from newyjht where ��ͬ���='" & frmYjBx.lblHtbh.Text & "' and ֧����=1"
'''''''''Set mod1.HTP = New ADODB.Recordset
'''''''''mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''Call frmYjBx.LrenH(frmYjBx.lblHtbh.Text)
'''''''''
'''''''''���÷�������е�����֧��
'''''''''ʵ�ʱ�
'''''''''tt = "Select sum(zFu) as zfu from yjz where htbh='" & frmYjBx.lblHtbh.Text & "'"
'''''''''mod1.HTT.Close
'''''''''mod1.HTT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''If IsNull(mod1.HTP.Fields(0).Value) = True Then
'''''''''    frmYjBx.lblCF.Caption = 0
'''''''''Else
'''''''''If IsNull(mod1.HTP.Fields(0).Value) = True Then
'''''''''    Ny = 0
'''''''''Else
'''''''''    Ny = mod1.HTP.Fields(0).Value
'''''''''End If
'''''''''    frmYjBx.lblCF.Caption = Ny + mod1.HTT.Fields("zfu").Value
'''''''''End If
'''''''''
'''''''''For oo = 0 To 6
'''''''''    frmYjBx.lblTm(oo).Caption = ""
'''''''''    frmYjBx.cmdQm(oo).Caption = ""
'''''''''    frmYjBx.lblQM(oo).Visible = False
'''''''''    frmYjBx.lblTm(oo).Visible = False
'''''''''    frmYjBx.cmdQm(oo).Visible = False
'''''''''Next
'''''''''
'''''''''�ж�����ǩ�ְ�ť , ��û��, �����
'''''''''If frmYjBx.lblYwy.Caption <> "" Then
'''''''''    tt = "select * from qmrz where btz=23 and qdbh='" & frmYjBx.lblYid.Caption & "' order by zid"
'''''''''    Set mod1.HTP = New ADODB.Recordset
'''''''''    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'''''''''
'''''''''    mod1.HTP.MoveFirst
'''''''''    For oo = 0 To 6
'''''''''        frmYjBx.lblQM(oo).Caption = mod1.HTP.Fields("qLabel").Value
'''''''''        If mod1.HTP.Fields("xf").Value = True Then
'''''''''            frmYjBx.cmdQm(oo).Caption = mod1.HTP.Fields("qren").Value
'''''''''            If frmYjBx.cmdQm(oo).Caption = "�Ͼ��쾭��" Then
'''''''''                frmYjBx.cmdQm(oo).Caption = "�Ͼ��쾭��"
'''''''''            End If
'''''''''            frmYjBx.lblTm(oo).Caption = mod1.HTP.Fields("qrq").Value
'''''''''        End If
'''''''''        frmYjBx.cmdQm(oo).Visible = True
'''''''''        frmYjBx.lblQM(oo).Visible = True
'''''''''        frmYjBx.lblTm(oo).Visible = True
'''''''''        mod1.HTP.MoveNext
'''''''''    Next
'''''''''    If frmYjBx.lblQM(5).Caption = "��֧��" Then
'''''''''        frmYjBx.lblQM(6).Visible = False
'''''''''        frmYjBx.cmdQm(6).Visible = False
'''''''''        frmYjBx.lblTm(6).Visible = False
'''''''''    End If
'''''''''    If Pwf = True And frmYjBx.cmdQm(5).Caption = "" And frmYjBx.cmdQm(6).Visible = False Then '��֧����ʾ
'''''''''        frmYjBx.cmdQm(5).Caption = frmYjBx.cmdQm(2).Caption
'''''''''        frmYjBx.lblTm(5).Caption = frmYjBx.lblTm(4).Caption
'''''''''    End If
'''''''''
'''''''''Else
'''''''''
'''''''''End If
'''''''''
'''''''''If QFF = False And mod1.DName = "�Ǽ���" And Pwf = True Then
'''''''''    frmYjBx.cmdWb.Visible = True
'''''''''Else
'''''''''    frmYjBx.cmdWb.Visible = False
'''''''''End If
'''''''''frmBxBrow.Enabled = False
'''''''''frmYjBx.Show
'''''''''frmYjBx.ZOrder 0
'''''''''frmYjBx.OptT1.Value = False
'''''''''frmYjBx.optT2.Value = False
Call frmYjBx.Bound(Yid)
Exit Sub
YZERR:
MsgBox "������ϣ�������һ�Σ���ر�ϵͳ������"
Set mod1.HTP = Nothing

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox Shift
'MsgBox KeyCode
If Shift = 2 And KeyCode = 76 Or KeyCode = 76 Then
    If mod1.Kyj = True Then
        If frmYj.Visible = False Then
            frmYj.Visible = True

        Else
            frmYj.Visible = False

        End If
    End If
    
End If
If Shift = 2 And KeyCode = 56 Then
    If Frame1.Visible = False Then
        Frame1.Visible = True
    Else
        Frame1.Visible = False
    End If
    
End If
End Sub

Private Sub Form_Load()
Set adoYj = New ADODB.Recordset
frmBxBrow.Width = mod1.FWidth
frmBxBrow.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
mga.Left = 50
mga.Top = 0
frmAdd.Top = -80
frmAdd.Left = frmBxBrow.Width - frmAdd.Width
'MGa.Height = frmBxBrow.Height - MGa.Top
mga.Height = frmBxBrow.Height - mga.Top - 1000
frmAdd.BorderStyle = 0
mtA.Value = mod1.DQda
Call GetWeek
dtgYJ.ColWidth(0) = 300
dtgYJ.ColWidth(3) = 2500
dtgYJ.ColWidth(5) = 2000
dtgYJ.ColWidth(1) = 0
dtgYJ.ColWidth(2) = 0
dtgYJ.ColWidth(8) = 0
dtgYJ.ColWidth(9) = 0
dtgYJ.ColWidth(10) = 0
dtgYJ.ColWidth(11) = 0
dtgYJ.ColWidth(12) = 0
dtgYJ.ColWidth(13) = 0
dtgYJ.ColWidth(15) = 0
dtgYJ.ColWidth(16) = 0
dtgYJ.ColWidth(17) = 0
dtgYJ.ColWidth(18) = 0
dtgYJ.ColWidth(19) = 0
dtgYJ.ColWidth(21) = 0
dtgYJ.ColWidth(22) = 0
'Call ResizeInit(Me) '�ڳ���װ��ʱ�������
'cmdBack.Left = Screen.Width - cmdBack.Width
'cmdBack.Top = Screen.Height - cmdBack.Height - 2000
If mod1.BmJl = False And (mod1.Bm = "ά����1" Or mod1.Bm = "ά����2" Or mod1.Bm = "��Ʒ��1" Or _
mod1.Bm = "��Ʒ��2" Or mod1.Bm = "�Ͼ���" Or mod1.Bm = "���ݰ�" Or mod1.Bm = "������" Or mod1.Bm = "�������۶���") And mod1.DName <> "���ĳ�" Then
    cmdFF.Visible = False
Else
    cmdFF.Visible = True
End If
Frame1.Visible = False
If mod1.KhK < 1 Then
    cmdFw.Visible = False
    lblFw.Visible = False
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
mga.Height = frmBxBrow.Height - mga.Top - 1000
frmAdd.Left = frmBxBrow.Width - frmAdd.Width
frmAdd.Height = frmBxBrow.Height - 1000
frmOpt.Top = mga.Height + mga.Top
'Call mod1.ResizeForm(Me) 'ȷ������ı�ʱ�ؼ���֮�ı�
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
Cancel = True
frmZu.TBa.Buttons(3).Value = tbrUnpressed
'frmBxBrow.WindowState = 0
frmBxBrow.Visible = False
End If
End Sub

Private Sub MGa_DblClick()

Static Px As Boolean

If mga.Row = 1 Then
    If Px = True Then
        mga.Sort = 2
        Px = False
    Else
        mga.Sort = 1
        Px = True
    End If
'Else
'    MsgBox MGa.ColData(1)
End If
'MGa.BackColorSel = vbGreen


End Sub


Private Sub mtA_Click()
    Call GetWeek
End Sub

Private Sub mtA_DateClick(ByVal DateClicked As Date)
    Call GetWeek
End Sub


Private Sub optMe_Click()
Dim tt As String
'Dim PK As String
On Error Resume Next

    tt = "FydV('" & mod1.DHid & "','" & mod1.DName & "')"
    frmBxBrow.AdoBxBro.Close
    frmBxBrow.AdoBxBro.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    Set frmBxBrow.mga.DataSource = frmBxBrow.AdoBxBro
    'PK = "<�� ʼ ��  |<�� �� ��  |>  �� �� |^ �� �� �� �� ��|> ǩ������ "
    'PK = "^�� �� �� Χ|^�� �� �� Χ|>  �� �� |^ �� �� �� �� ��|> ǩ������ "
    'frmBxBrow.mga.FormatString = PK
End Sub

Private Sub optQi_Click()
Dim tt As String
Dim pk As String
On Error Resume Next


        tt = "CBXNew('" & mod1.DHid & "')"
        frmBxBrow.AdoBxBro.Close
        frmBxBrow.AdoBxBro.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        Set frmBxBrow.mga.DataSource = frmBxBrow.AdoBxBro



End Sub

Public Sub GetWeek()
Select Case mtA.DayOfWeek
Case 1 '������
lblFr.Caption = DateSerial(Year(mtA.Value), Month(mtA.Value), Day(mtA.Value) - 6)
lblLr.Caption = mtA.Value
Case 2 '����һ
lblFr.Caption = mtA.Value
lblLr.Caption = DateSerial(Year(mtA.Value), Month(mtA.Value), Day(mtA.Value) + 6)
Case 3
lblFr.Caption = DateSerial(Year(mtA.Value), Month(mtA.Value), Day(mtA.Value) - 1)
lblLr.Caption = DateSerial(Year(mtA.Value), Month(mtA.Value), Day(mtA.Value) + 5)
Case 4
lblFr.Caption = DateSerial(Year(mtA.Value), Month(mtA.Value), Day(mtA.Value) - 2)
lblLr.Caption = DateSerial(Year(mtA.Value), Month(mtA.Value), Day(mtA.Value) + 4)
Case 5
lblFr.Caption = DateSerial(Year(mtA.Value), Month(mtA.Value), Day(mtA.Value) - 3)
lblLr.Caption = DateSerial(Year(mtA.Value), Month(mtA.Value), Day(mtA.Value) + 3)
Case 6
lblFr.Caption = DateSerial(Year(mtA.Value), Month(mtA.Value), Day(mtA.Value) - 4)
lblLr.Caption = DateSerial(Year(mtA.Value), Month(mtA.Value), Day(mtA.Value) + 2)
Case 7
lblFr.Caption = DateSerial(Year(mtA.Value), Month(mtA.Value), Day(mtA.Value) - 5)
lblLr.Caption = DateSerial(Year(mtA.Value), Month(mtA.Value), Day(mtA.Value) + 1)
End Select
End Sub
