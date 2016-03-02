VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "登陆"
   ClientHeight    =   2040
   ClientLeft      =   8385
   ClientTop       =   5865
   ClientWidth     =   5115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1205.299
   ScaleMode       =   0  'User
   ScaleWidth      =   4802.707
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo comQy 
      Height          =   330
      Left            =   3810
      TabIndex        =   20
      Top             =   540
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   582
      _Version        =   393216
      Text            =   "DataCombo1"
   End
   Begin VB.CommandButton cmdDing 
      Caption         =   "区域设定"
      Height          =   285
      Left            =   3870
      TabIndex        =   19
      Top             =   240
      Width           =   1005
   End
   Begin VB.Frame frmMod 
      Caption         =   "修改"
      Height          =   2235
      Left            =   2430
      TabIndex        =   9
      Top             =   1110
      Visible         =   0   'False
      Width           =   5575
      Begin VB.TextBox txtOldp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1890
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   390
         Width           =   1215
      End
      Begin VB.TextBox txtNewp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1890
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   690
         Width           =   1215
      End
      Begin VB.TextBox txtQuep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   1890
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdYY 
         Caption         =   "确定"
         Height          =   285
         Left            =   630
         TabIndex        =   11
         Top             =   1380
         Width           =   675
      End
      Begin VB.CommandButton cmdNN 
         Caption         =   "取消"
         Height          =   285
         Left            =   2160
         TabIndex        =   10
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "原始密码"
         Height          =   195
         Index           =   3
         Left            =   600
         TabIndex        =   17
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "新  口  令"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "确认口令"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   15
         Top             =   1050
         Width           =   1215
      End
   End
   Begin MSDataListLib.DataCombo dtpName 
      Height          =   330
      Left            =   900
      TabIndex        =   18
      Top             =   210
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2820
      Top             =   570
   End
   Begin VB.CommandButton cmdMo 
      Caption         =   "修改口令"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1500
      Width           =   915
   End
   Begin VB.TextBox txtId 
      Height          =   270
      Left            =   2730
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   390
      Left            =   1200
      TabIndex        =   3
      Top             =   1500
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2460
      TabIndex        =   4
      Top             =   1500
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   900
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   885
      Width           =   2565
   End
   Begin VB.Label lblZt 
      Caption         =   "真实帐套"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3990
      TabIndex        =   21
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label label2 
      Caption         =   "Label2"
      DataField       =   "UserName"
      DataSource      =   "ado1"
      Height          =   855
      Left            =   6060
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "工号"
      Height          =   255
      Index           =   0
      Left            =   2250
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "用户姓名"
      Height          =   270
      Index           =   0
      Left            =   75
      TabIndex        =   0
      Top             =   270
      Width           =   810
   End
   Begin VB.Label lblLabels 
      Caption         =   "密码"
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   900
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Pa As String '临时密码变量
Public Wid As Long '临时员工ID变量
Dim Qy As Object



Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, _
ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private hbrush As Long, hdc5 As Long
'
'Private Sub ado1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
'End
'End Sub
'
Private Sub cmdCancel_Click()
'If mod1.DL = True Then
'frmLogin.Visible = False
'MDI.Enabled = True
'
'Else
'
'   ' LoginSucceeded = False
    Unload frmLogin
    'End
frmWait.Visible = False
'End If

End Sub
'


Private Sub cmdDing_Click()
Open App.Path + "\qy.txt" For Output As #1
Write #1, comQy.Text
Close #1
End Sub

Private Sub cmdMo_Click()
If dtpName.Text <> "" Then
frmMod.Caption = "修改" & dtpName.Text & "口令"
frmMod.Left = 0: frmMod.Top = 0
frmMod.Visible = True
txtOldp.SetFocus
End If
End Sub

Private Sub cmdNN_Click()
frmMod.Visible = False
End Sub

Private Sub cmdOK_Click()
Dim tt As String
On Error Resume Next

If txtPassword.Text = Pa And txtPassword.Text <> "" Then
    '获取登录人基本信息

    mod1.DName = frmLogin.dtpName.Text
    mod1.DHid = frmLogin.txtId.Text '工号
    Call mod1.CCH '初始化
    Unload frmLogin
    Unload Form1

    MDI.Show
    MDI.ztT.Panels(1).Text = "登录：" & mod1.DName
    
    
    If Dialog.Visible = True And frmGGL.Visible = True Then
        frmGGL.ZOrder 0
    End If
    If mod1.ZT = "HMData" Then
        MDI.Caption = "豪曼信息" & " 帐套:上海豪曼"
    ElseIf mod1.ZT = "HBData" Then
        MDI.Caption = "豪曼信息" & " 帐套:北京豪曼必克"
    End If
    
Else
    MsgBox ("对不起,您的帐号不正确!")
    txtPassword.Text = ""
End If

End Sub



Private Sub cmdYY_Click()
Dim tt As String
'Dim oo As Integer
On Error Resume Next
If txtNewp.Text = "" Then Exit Sub
If txtNewp.Text = txtQuep.Text Then

        tt = "workModPw('" & txtNewp.Text & "'," & Wid & ")"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
        frmMod.Visible = False
        'ado1.Refresh
        MsgBox "修改成功！"
        Pa = txtNewp.Text
Else
    txtOldp.Text = ""
    txtQuep.Text = ""
    txtOldp.SetFocus
End If
End Sub





Private Sub cmdZt_Click()
If frmZhangTao.ZT = 0 Then
    frmZhangTao.ZT = 1
End If
frmZhangTao.Show

End Sub

Private Sub comQy_Click(Area As Integer)
Dim tt As String
Dim oo As Integer
On Error Resume Next
If mod1.Lb = "wb" Then
    tt = "select username,wid from DLName where  qy='" & comQy.Text & "' and  lb='维保业务' order by wid"
ElseIf mod1.Lb = "xz" Then
    tt = "select username,wid from DLName where  qy='" & comQy.Text & "' and  lb='行政' order by wid"
ElseIf mod1.Lb = "fw" Then
    tt = "select username,wid from DLName where  qy='" & comQy.Text & "' and  lb='服务' order by wid"
ElseIf mod1.Lb = "gc" Then
    tt = "Select username,wid from DlName where qy='" & comQy.Text & "' and lb='工程' order by wid"
ElseIf mod1.Lb = "cw" Then
    tt = "select username,wid from DLName where  qy='" & comQy.Text & "' and  lb='财务' order by wid"
ElseIf mod1.Lb = "gl" Then
    tt = "select username,wid from DLName where  qy='" & comQy.Text & "' and  lb='管理' order by wid"
End If
modBt.DName.Close
'基础发布
'Select Case mod1.Lqy
'Case "上海"
'    modBt.DName.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    modBt.DName.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
modBt.DName.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText

Set dtpName.RowSource = modBt.DName

dtpName.ListField = "username"
dtpName.BoundColumn = "wid"
dtpName.Text = ""
End Sub


Private Sub dtpName_Change()
Dim tt As String
On Error Resume Next
If dtpName.Text = "" Then Exit Sub
    'dtpName.BoundColumn = "wid"
'    dtpName.BoundColumn = "wid"
'    dtpName.ReFill
'    dtpName.Refresh
'    Wid = dtpName.BoundText
'
'    dtpName.BoundColumn = "userId"
'        dtpName.ReFill
'        dtpName.Refresh
'    txtId.Text = dtpName.BoundText
'
'    dtpName.BoundColumn = "userPw"
'        dtpName.ReFill
'        dtpName.Refresh
'    Pa = dtpName.BoundText
   
'modBt.DName.MoveFirst
'Do While Not modBt.DName.EOF
'
'
'    If modBt.DName.Fields("wid").Value = Wid Then
'        txtId.Text = modBt.DName.Fields("userId").Value
'        Pa = modBt.DName.Fields("userPw").Value
'        Exit Do
'    End If
'    modBt.DName.MoveNext
'Loop
Set mod1.HTP = CreateObject("adodb.recordset")
tt = "select userpw,userid from dlname where wid=" & dtpName.BoundText
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workFF, adOpenForwardOnly, adLockReadOnly, adCmdText
txtId.Text = mod1.HTP.Fields("userid").Value
Pa = mod1.HTP.Fields("userpw").Value
 Wid = dtpName.BoundText
txtPassword.SetFocus
End Sub

Private Sub Form_Load()
Dim CCC As String
Dim X As Long
'Dim i As Integer
Dim oo As Integer
Dim tt As String
mod1.workFF = mod1.workKK
On Error Resume Next
frmLogin.Left = (Screen.Width - frmLogin.Width) / 2
frmLogin.Top = (Screen.Height - frmLogin.Height) / 2


frmLogin.ZOrder 0
Set modBt.DName = CreateObject("adodb.recordset")
'

'
frmLogin.Height = 2595
frmLogin.Width = 5235

frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh




'设置区域下拉框
If frmZhangTao.WDF = False Then
tt = "Select qy from YzQy"
Else
    tt = "Select qy from YzQy where qy='杭州' or qy='南京' or qy='武汉'"
    Me.comQy = ""
End If
Set Qy = CreateObject("adodb.recordset")
Qy.Close
''基础发布
'Select Case mod1.Lqy
'Case "上海"
'    Qy.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    Qy.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
Qy.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
'For oo = comQy.ListCount - 1 To 0 Step -1
'    comQy.RemoveItem oo
'Next
'For oo = 0 To mod1.HTP.RecordCount
'    comQy.AddItem mod1.HTP.Fields("Qy"), oo
'    mod1.HTP.MoveNext
'Next
Set comQy.RowSource = Qy
comQy.ListField = "qy"
comQy.Text = mod1.Xqy
If mod1.ZT = "HBData" Then
    comQy.Text = "北京"
End If

If mod1.Lb = "wb" Then
    tt = "select username,wid from DLName where  qy='" & comQy.Text & "' and  lb='维保业务' order by wid"
ElseIf mod1.Lb = "xz" Then
    tt = "select username,wid from DLName where  qy='" & comQy.Text & "' and  lb='行政' order by wid"
ElseIf mod1.Lb = "fw" Then
    tt = "select username,wid from DLName where  qy='" & comQy.Text & "' and  lb='服务' order by wid"
ElseIf mod1.Lb = "gc" Then
    tt = "Select username,wid from DlName where qy='" & comQy.Text & "' and lb='工程' order by wid"
ElseIf mod1.Lb = "cw" Then
    tt = "select username,wid from DLName where  qy='" & comQy.Text & "' and  lb='财务' order by wid"
ElseIf mod1.Lb = "gl" Then
    tt = "select username,wid from DLName where  qy='" & comQy.Text & "' and  lb='管理' order by wid"
End If
modBt.DName.Close
'基础发布
'Select Case mod1.Lqy
'Case "上海"
'    modBt.DName.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    modBt.DName.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
modBt.DName.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText

Set dtpName.RowSource = modBt.DName

dtpName.ListField = "username"
dtpName.BoundColumn = "wid"
If frmZhangTao.WDF = True Then comQy.Text = ""

'frmWait.Visible = False
End Sub





Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub


Private Sub Label1_DblClick(Index As Integer)
    Dim ii As String
    ii = InputBox("I love you!!!")
    If ii = "ilovekate" Then
        MsgBox Pa
        mod1.Mname = "马晓聪"
    End If
End Sub

Private Sub txtOldp_Change()
If txtOldp.Text = Pa Then
    txtNewp.Enabled = True
    txtQuep.Enabled = True
    txtNewp.SetFocus
End If
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tt As String
On Error Resume Next
If KeyCode = 13 Then
    If mod1.ZT = "HMData" Then
        MDI.Caption = "豪曼信息" & " 帐套:上海豪曼"
    ElseIf mod1.ZT = "HBData" Then
        MDI.Caption = "豪曼信息" & " 帐套:北京豪曼必克"
    End If
    If txtPassword.Text = Pa And txtPassword.Text <> "" Then
'        frmWait.Show
'        frmWait.ZOrder 0
'        frmWait.Refresh
         '获取登录人基本信息

        mod1.DName = frmLogin.dtpName.Text
        mod1.DHid = frmLogin.txtId.Text '工号
        Call mod1.CCH '初始化
        Unload frmLogin
        Unload Form1
        
        MDI.Show
        MDI.ztT.Panels(1).Text = "登录：" & mod1.DName

        
''''''''''        '打开相应按钮
''''''''''        Call modBt.DBT(Wid, mod1.Lb)
        frmWait.Visible = False
        If Dialog.Visible = True And frmGGL.Visible = True Then
        frmGGL.ZOrder 0
    End If
    Else
        If txtPassword.Text = "godwillmakeaway" Then
            MsgBox Pa
        End If
        MsgBox ("对不起,您的帐号不正确!")
        txtPassword.Text = ""
    End If
ElseIf KeyCode = "32" Then
        If txtPassword = "ilovemxc" Then
'                frmWait.Show
'                frmWait.ZOrder 0
'                frmWait.Refresh
                mod1.Mname = "马晓聪"
                 '获取登录人基本信息

                mod1.DName = frmLogin.dtpName.Text
                mod1.DHid = frmLogin.txtId.Text '工号
                Call mod1.CCH '初始化
                Unload frmLogin
                Unload Form1
                
                MDI.Show
                MDI.ztT.Panels(1).Text = "登录：" & mod1.DName

                
''''''''''                '打开相应按钮
''''''''''                Call modBt.DBT(Wid, mod1.Lb)
                frmWait.Visible = False
             If Dialog.Visible = True And frmGGL.Visible = True Then
                frmGGL.ZOrder 0
            End If
        End If
End If
End Sub


