VERSION 5.00
Begin VB.Form HLB 
   BackColor       =   &H00C0FFC0&
   Caption         =   "���ܲ�"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10740
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   9510
   ScaleWidth      =   10740
   StartUpPosition =   3  '����ȱʡ
   Visible         =   0   'False
   Begin VB.CommandButton cmdDel 
      Caption         =   "ɾ��"
      Height          =   405
      Left            =   8280
      TabIndex        =   34
      Top             =   9060
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "���"
      Height          =   405
      Left            =   7770
      TabIndex        =   33
      Top             =   9060
      Width           =   495
   End
   Begin VB.CommandButton cmdMod 
      Height          =   405
      Left            =   8790
      Picture         =   "HLB.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "�޸�"
      Top             =   9060
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   405
      Left            =   9270
      Picture         =   "HLB.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "����"
      Top             =   9060
      Width           =   465
   End
   Begin VB.TextBox txtH 
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   10
      Left            =   7200
      TabIndex        =   29
      Text            =   "10"
      Top             =   6330
      Width           =   2475
   End
   Begin VB.TextBox txtH 
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   4
      Left            =   6870
      TabIndex        =   27
      Text            =   "4"
      Top             =   1410
      Width           =   2715
   End
   Begin VB.TextBox txtH 
      ForeColor       =   &H00C000C0&
      Height          =   285
      Index           =   3
      Left            =   2310
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "3"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtH 
      ForeColor       =   &H00C000C0&
      Height          =   315
      Index           =   2
      Left            =   6870
      TabIndex        =   25
      Text            =   "2"
      Top             =   810
      Width           =   2715
   End
   Begin VB.TextBox txtH 
      ForeColor       =   &H00C000C0&
      Height          =   270
      Index           =   1
      Left            =   2310
      TabIndex        =   24
      Text            =   "1"
      Top             =   870
      Width           =   2295
   End
   Begin VB.TextBox txtH 
      Height          =   345
      Index           =   14
      Left            =   7200
      TabIndex        =   23
      Text            =   "14"
      Top             =   8190
      Width           =   2475
   End
   Begin VB.TextBox txtH 
      Height          =   345
      Index           =   12
      Left            =   7200
      TabIndex        =   22
      Text            =   "12"
      Top             =   7260
      Width           =   2475
   End
   Begin VB.TextBox txtH 
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   13
      Left            =   2280
      TabIndex        =   21
      Text            =   "13"
      Top             =   8190
      Width           =   2865
   End
   Begin VB.TextBox txtH 
      ForeColor       =   &H00FF0000&
      Height          =   345
      Index           =   11
      Left            =   2280
      TabIndex        =   20
      Text            =   "11"
      Top             =   7290
      Width           =   2865
   End
   Begin VB.TextBox txtH 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   9
      Left            =   3810
      TabIndex        =   18
      Text            =   "9"
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox txtH 
      ForeColor       =   &H000000FF&
      Height          =   855
      Index           =   7
      Left            =   2310
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Tag             =   "100"
      Text            =   "HLB.frx":0974
      Top             =   4980
      Width           =   7395
   End
   Begin VB.TextBox txtH 
      ForeColor       =   &H00FF0000&
      Height          =   765
      Index           =   6
      Left            =   2310
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Tag             =   "200"
      Text            =   "HLB.frx":0976
      Top             =   3960
      Width           =   7395
   End
   Begin VB.TextBox txtH 
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   8
      Left            =   2280
      TabIndex        =   11
      Text            =   "8"
      Top             =   6360
      Width           =   1035
   End
   Begin VB.TextBox txtH 
      ForeColor       =   &H00C000C0&
      Height          =   1695
      Index           =   5
      Left            =   2310
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "300"
      Text            =   "HLB.frx":0978
      Top             =   2070
      Width           =   7365
   End
   Begin VB.Label lblAl 
      BackStyle       =   0  'Transparent
      Caption         =   "------  ���䰸��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   5490
      TabIndex        =   36
      Top             =   120
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Label lblLc 
      Caption         =   "lblLc"
      Height          =   315
      Left            =   2190
      TabIndex        =   35
      Top             =   8880
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblHid 
      Caption         =   "lblHid"
      Height          =   225
      Left            =   3840
      TabIndex        =   32
      Top             =   9000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "�ܽ���ǩ�֣�"
      Height          =   225
      Left            =   5550
      TabIndex        =   28
      Top             =   6360
      Width           =   1275
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   6360
      Width           =   315
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "���ڣ�"
      Height          =   285
      Left            =   6270
      TabIndex        =   15
      Top             =   8190
      Width           =   615
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "�ܾ���������"
      Height          =   285
      Left            =   5760
      TabIndex        =   14
      Top             =   7350
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "���ڣ�"
      Height          =   285
      Left            =   900
      TabIndex        =   13
      Top             =   8220
      Width           =   1125
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "���ܲ�ίԱ��ǩ�£�"
      Height          =   285
      Left            =   330
      TabIndex        =   12
      Top             =   7350
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   285
      Left            =   3420
      TabIndex        =   10
      Top             =   6360
      Width           =   315
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "�������ˣ����ţ�"
      Height          =   285
      Left            =   450
      TabIndex        =   9
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "�ܽ��˸���"
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   5070
      Width           =   1425
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "������ʩ"
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   4020
      Width           =   1905
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "�ܽ���Ϊ�����������"
      Height          =   285
      Left            =   -90
      TabIndex        =   5
      Top             =   2160
      Width           =   2115
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Top             =   1470
      Width           =   885
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   285
      Left            =   570
      TabIndex        =   3
      Top             =   1470
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "�����ˣ����ţ�"
      Height          =   285
      Left            =   5160
      TabIndex        =   2
      Top             =   900
      Width           =   1755
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "����ʱ��"
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   900
      Width           =   1665
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   4140
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "HLB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Sub HLBQing()
Dim oo As Integer
For oo = 1 To 14
    txtH(oo).Text = ""
Next
lblHid.Caption = ""
lblLc.Caption = ""
lblAl.Visible = False
End Sub

Private Sub cmdAdd_Click()
Dim Err As String
On Error Resume Next
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "HLBAdd"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@t2") = mod1.DName
mod1.cmd.Execute

Err = mod1.cmd.Parameters("@errinf").Value

If Err <> "" Then
    MsgBox "ϵͳ���ϣ�����ִ��һ�Σ�"
Else
    Call HLBQing
    lblHid.Caption = mod1.cmd.Parameters("@hid").Value
    txtH(1).Text = mod1.cmd.Parameters("@t1").Value
    txtH(2).Text = mod1.DName
    lblLc.Caption = 1
    cmdSave.Enabled = True
    cmdDel.Enabled = True
    lblAl.Visible = False
End If
Set cmd = Nothing
End Sub

Private Sub cmdDel_Click()
Dim Err As String
On Error Resume Next

Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "HLBTT"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@errinf").Value = ""
mod1.cmd.Execute
Err = mod1.cmd.Parameters("@errinf").Value
Set cmd = Nothing
If Err <> "�ɹ�" Then
    MsgBox "ϵͳ���ϣ�����ִ��һ�Σ�"
    Exit Sub
ElseIf Err = "�ɹ�" Then
    MsgBox "ok"
End If
End Sub

Private Sub cmdSave_Click()
Dim Err As String
On Error Resume Next
txtH(8).Text = txtH(2).Text
If lblLc.Caption = 1 Then
    If txtH(3).Text = "" Or txtH(4).Text = "" Or txtH(5).Text = "" Then
        ii = MsgBox("������δд�������Ľ����ı�Ҫ��Ϣ��", vbOKOnly + vbInformation, "��ȷ��")
        Exit Sub
    End If
ElseIf lblLc.Caption = 2 Then
    If txtH(6).Text = "" Or txtH(9).Text = "" Or txtH(11).Text = "" Then
        ii = MsgBox("���ܲ�ίԱ���Աδд��������ı�Ҫ��Ϣ��", vbOKOnly + vbInformation, "��ȷ��")
        Exit Sub
    End If
ElseIf lblLc.Caption = 3 Then
    If txtH(10).Text = "" Then
        ii = MsgBox("�ܽ���δǩ�֣�", vbOKOnly + vbInformation, "��ȷ��")
        txtH(10).SetFocus
        Exit Sub
    End If
Else
    Exit Sub
End If
lblLc.Caption = lblLc.Caption + 1
Set mod1.cmd = CreateObject("adodb.command")
mod1.cmd.ActiveConnection = mod1.cc
mod1.cmd.CommandText = "HLBsave"
mod1.cmd.CommandType = adCmdStoredProc
mod1.cmd.Parameters("@t1") = txtH(1).Text
mod1.cmd.Parameters("@t2") = txtH(2).Text
mod1.cmd.Parameters("@t3") = txtH(3).Text
mod1.cmd.Parameters("@t4") = txtH(4).Text
mod1.cmd.Parameters("@t5") = txtH(5).Text
mod1.cmd.Parameters("@t6") = txtH(6).Text
mod1.cmd.Parameters("@t7") = txtH(7).Text
mod1.cmd.Parameters("@t8") = txtH(8).Text
mod1.cmd.Parameters("@t9") = Val(txtH(9).Text)
mod1.cmd.Parameters("@t10") = txtH(10).Text
mod1.cmd.Parameters("@t11") = txtH(11).Text
mod1.cmd.Parameters("@t12") = txtH(12).Text
mod1.cmd.Parameters("@t13") = Date
mod1.cmd.Parameters("@t14") = Date
mod1.cmd.Parameters("@lc") = lblLc.Caption
mod1.cmd.Parameters("@hid") = lblHid.Caption
mod1.cmd.Execute
Err = mod1.cmd.Parameters("@errinf").Value
Set cmd = Nothing
If Err <> "�ɹ�" Then
    MsgBox "ϵͳ���ϣ�����ִ��һ�Σ�"
    lblLc.Caption = lblLc.Caption - 1
    Exit Sub
ElseIf Err = "�ɹ�" Then
    If lblLc.Caption = 2 Then
        MsgBox "���Ľ��������Ѿ��ύ�����ܲ�ίԱ�Ὣ��������������飬лл��"
        frmGGL.Visible = False
    ElseIf lblLc.Caption = 3 Then
        MsgBox "�˽�������Ч����ת���ܽ���" & txtH(3).Text & "��ȷ�ϣ�"
    ElseIf lblLc.Caption = 4 Then
        MsgBox "��ϲ�����Ĺ����¼����������£� ����"
    End If
    cmdSave.Enabled = False
    cmdDel.Enabled = False
End If

End Sub


Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim ii As Integer
Dim tt As String
On Error Resume Next
If txtH(3).Text <> "С����" Then
    If lblLc.Caption = 1 And cmdSave.Enabled = True And cmdSave.Visible = True Then
        ii = MsgBox("�����˳�������������뵥�ӣ�", vbYesNo + vbInformation, "��ȷ��")
        If ii = vbYes Then
            tt = "delete from HLB where hid=" & lblHid.Caption
            Set mod1.HTP = CreateObject("adodb.recordset")
            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
            frmGGL.frmLx.Enabled = False
            frmGGL.cmdZX.Enabled = True
            frmGGL.cmdSave.Enabled = False
            frmGGL.cmdReply.Enabled = False
        Else
            Cancel = True
        End If
    End If
Else
    frmGGL.frmLx.Enabled = False
    frmGGL.cmdZX.Enabled = True
    frmGGL.cmdSave.Enabled = False
    frmGGL.cmdReply.Enabled = False
End If
If Dialog.Visible = True Then
    Dialog.Enabled = True
Else
    frmZu.Enabled = True
End If
End Sub


Private Sub txtH_Change(Index As Integer)
If Index >= 5 And Index <= 7 Then
If Len(txtH(Index).Text) >= Val(txtH(Index).Tag) Then
    MsgBox ("���������ﵽ" & txtH(Index).Tag & ",�������ƽ����ܱ��棡")
    
End If
End If
End Sub

Private Sub txtH_DblClick(Index As Integer)
If Index = 11 Then
    txtH(11).Text = mod1.DName
    txtH(13).Text = mod1.DQda
ElseIf Index = 10 Then
    txtH(10).Text = mod1.DName
ElseIf Index = 3 Then
    Set Ren.XForm = New HLB
    Call mod1.RenXz("HLB", Me, 0)
End If
End Sub



Public Sub HLBLI()
txtH(1).Text = Date
txtH(2).Text = "Сè����"
txtH(3).Text = "С����"
txtH(4).Text = "10"
txtH(5).Text = "    ��������ɭ����,���������˴���ǡ�С����Ϊ�˾���,�����Լ������ǿڵ�Σ�գ������侲�������¸�ͬ���������,������ľͷǹ�����˴���ǡ�Ϊ�˸�лС���ã�������ܲ�ίԱ��������룡"
txtH(6).Text = "    �������ܲ�ίԱ�����������о����϶�Сè����˵�������ȫ��ʵ�����ܲ�ίԱ��Ϊ����С����������Σ���塢�Ἲ���˵ľ����ؽ�������Ե��ش��������ں��ܲ�һ�����������������������ֵ���ʹ���Ժ���������ܴ���ǡ�"
txtH(7).Text = "    ������ֵ�����̫�����ˣ��еĳԣ����е��棡 ����"
txtH(8).Text = "Сè����"
txtH(9).Text = "2"
txtH(10).Text = "С����"
txtH(11).Text = "��è����"
txtH(12).Text = "�ϻ�����ķ"
txtH(13).Text = Date
txtH(14).Text = Date
lblAl.Visible = True
End Sub

Public Sub HLBBound(Hid As Long)
Dim oo As Integer
Dim tt As String
On Error Resume Next
tt = "select * from Hlb where hid=" & Hid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
For oo = 1 To 14
    txtH(oo).Text = mod1.HTP.Fields("t" & oo).Value
Next
lblLc.Caption = mod1.HTP.Fields("lc").Value
lblHid.Caption = Hid
End Sub
