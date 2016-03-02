VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.MDIForm MDI 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808080&
   Caption         =   "豪曼信息"
   ClientHeight    =   10650
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   10230
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer timFl 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   3810
      Top             =   3450
   End
   Begin VB.Timer timXz 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   1680
   End
   Begin MSAdodcLib.Adodc adoMaxG 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      Top             =   1005
      Visible         =   0   'False
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\work.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\work.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "worker"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar ztT 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   10185
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   21167
            MinWidth        =   21167
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1005
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   10170
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10230
      Begin MSComCtl2.DTPicker dtpTemp 
         Height          =   225
         Left            =   5520
         TabIndex        =   4
         Top             =   90
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   397
         _Version        =   393216
         Format          =   134610945
         CurrentDate     =   38776
      End
      Begin VB.TextBox Text1 
         DataField       =   "UserId"
         DataSource      =   "adoMa"
         Height          =   345
         Left            =   6180
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   390
         Width           =   2385
      End
      Begin MSAdodcLib.Adodc adoMa 
         Height          =   765
         Left            =   1050
         Top             =   90
         Visible         =   0   'False
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   1349
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\HMXP9000\work.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\HMXP9000\work.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "worker"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         DataField       =   "UserId"
         DataSource      =   "adoMaxG"
         Height          =   315
         Left            =   9060
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Menu Mfile 
      Caption         =   "文件"
      WindowList      =   -1  'True
      Begin VB.Menu Clogin 
         Caption         =   "重新登录"
         Enabled         =   0   'False
      End
      Begin VB.Menu Mexit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu Medit 
      Caption         =   "编辑"
   End
   Begin VB.Menu Mwindow 
      Caption         =   "窗口"
      Begin VB.Menu Mywdht 
         Caption         =   "业务导航图"
      End
   End
   Begin VB.Menu Base 
      Caption         =   "基础设置"
      Begin VB.Menu MFY 
         Caption         =   "费用"
      End
      Begin VB.Menu xZ 
         Caption         =   "行政"
         Begin VB.Menu GSZL 
            Caption         =   "公司资料"
         End
         Begin VB.Menu JZ 
            Caption         =   "合同基准"
         End
         Begin VB.Menu RS 
            Caption         =   "人事档案"
         End
      End
      Begin VB.Menu yZ 
         Caption         =   "运作"
         Begin VB.Menu qyHf 
            Caption         =   "区域划分"
         End
         Begin VB.Menu fkFC 
            Caption         =   "付款方式"
         End
      End
      Begin VB.Menu XS 
         Caption         =   "销售"
         Begin VB.Menu hyXz 
            Caption         =   "客户行业"
         End
      End
      Begin VB.Menu GC 
         Caption         =   "工程"
         Begin VB.Menu JZPB 
            Caption         =   "机组品牌"
         End
         Begin VB.Menu ZCG 
            Caption         =   "正常工时"
         End
         Begin VB.Menu jjR 
            Caption         =   "国定假日"
         End
         Begin VB.Menu LT 
            Caption         =   "工时费用"
         End
      End
      Begin VB.Menu adMIN 
         Caption         =   "系统管理员"
      End
   End
   Begin VB.Menu XiXi 
      Caption         =   "消息"
      Begin VB.Menu xa 
         Caption         =   "sales1"
      End
   End
   Begin VB.Menu BZ 
      Caption         =   "帮助"
      Begin VB.Menu GY 
         Caption         =   "关于"
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Cq As Boolean '按面板的返回时将unload MDI窗体而不出现询问退出框。

Dim adoG As Object '探测有无新消息的ado

  Const FLASHW_STOP = 0                                                                   'Stop   flashing.   The   system   restores _
                                                                                                          the   window   to   its   original   state.
  Const FLASHW_CAPTION = &H1                                                         'Flash   the   window   caption.
  Const FLASHW_TRAY = &H2                                                               'Flash   the   taskbar   button.
  Const FLASHW_ALL = (FLASHW_CAPTION Or FLASHW_TRAY)             'Flash   both   the   window   caption   and   taskbar   button.   This   is _
                                                                                                          equivalent   to   setting   the   FLASHW_CAPTION   Or   FLASHW_TRAY   flags.
  Const FLASHW_TIMER = &H4                                                             'Flash   continuously,   until   the   FLASHW_STOP   flag   is   set.
  Const FLASHW_TIMERNOFG = &HC                                                     'Flash   continuously   until   the   window   comes   to   the   foreground.
  Private Type FLASHWINFO
      cbSize         As Long
      hwnd             As Long
      dwFlags       As Long
      uCount         As Long
      dwTimeout   As Long
  End Type
  Private Declare Function FlashWindowEx Lib "user32" (pfwi As FLASHWINFO) As Boolean
  Private Declare Sub Sleep Lib "kernel32" _
   (ByVal dwMilliseconds As Long)
   

   





Private Sub JZ_Click()
Dim tt As String
Dim Ra, Rb
Dim La, Lb
Dim oo As Integer
If mod1.DHid = "HM003" Or mod1.DName = "宋晓炯1" Or mod1.DName = "马晓聪" Or mod1.DName = "周春云" Then
On Error Resume Next
frmJiZun.Show
tt = "select jz from jizun order by jid;" & _
    "select jz from jizunFk order by jid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
frmJiZun.txt1.Text = Format(Val(Ra(0, 0)), "0.00")
frmJiZun.txt2.Text = Format(Val(Ra(0, 1)), "0.00")
frmJiZun.txt3.Text = Format(Val(Ra(0, 2)), "0.00")
frmJiZun.txtA.Text = Format(Val(Ra(0, 4)), "0.00")
frmJiZun.txt6.Text = Format(Val(Ra(0, 5)), "0.00")
frmJiZun.txtB.Text = Format(Val(Ra(0, 6)), "0.00")
frmJiZun.txtC.Text = Format(Val(Ra(0, 7)), "0.00")
frmJiZun.txtD.Text = Format(Val(Ra(0, 8)), "0.00")
frmJiZun.txtE.Text = Format(Val(Ra(0, 9)), "0.00")
frmJiZun.txtF.Text = Format(Val(Ra(0, 10)), "0.00")
For oo = 0 To 8
    frmJiZun.txtCW(oo).Text = Format(Val(Rb(0, oo)), "0.00")
Next
'''''frmJiZun.txtA1.Text = Format(Val(Rb(0, 0)), "0.00")
'''''frmJiZun.txtA2.Text = Format(Val(Rb(0, 1)), "0.00")
'''''frmJiZun.txtA3.Text = Format(Val(Rb(0, 2)), "0.00")
'''''frmJiZun.txtB1.Text = Format(Val(Rb(0, 3)), "0.00")
'''''frmJiZun.txtB2.Text = Format(Val(Rb(0, 4)), "0.00")
'''''frmJiZun.txtB3.Text = Format(Val(Rb(0, 5)), "0.00")
'''''frmJiZun.txtB4.Text = Format(Val(Rb(0, 6)), "0.00")
'''''frmJiZun.txtB5.Text = Format(Val(Rb(0, 7)), "0.00")
End If
End Sub

Private Sub MDIForm_Activate()
timFl.Enabled = False
End Sub

Private Sub MDIForm_Click()
timFl.Enabled = False
End Sub

Private Sub MDIForm_DblClick()
frmAbout.Show
End Sub

Private Sub MDIForm_Load()
Set adoG = CreateObject("adodb.recordset")
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'''''''''Dim aa As Integer
'''''''''Dim tt As String
'''''''''On Error Resume Next
'''''''''
'''''''''    aa = MsgBox("是否真的要退出豪曼系统", vbInformation + vbYesNo)
'''''''''    If aa = vbYes Then
'''''''''    '    '将"正使用者"字段清空,以使其他人可用
'''''''''    '    tt = "update htping set DKR='' where DKR='" & frmLogin.Combo1.Text & "'"
'''''''''    '    Set mod1.HTP = CreateObject("adodb.recordset")
'''''''''    '    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'''''''''    '    Call mod1.DelDKZ  '退出表单时删除打开记录,以让别人能打开此单据
'''''''''    '    'Call mod1.delTmpB '删除临时表
'''''''''        Call mod1.zhuLK '退出时取消注册
'''''''''
'''''''''        'MsgBox "OK"
'''''''''        Unload frmOL
'''''''''        Unload frmGGL
'''''''''        End
'''''''''    Else
'''''''''        Cancel = True
'''''''''        frmZu.Visible = True
'''''''''        frmZu.ZOrder 0
'''''''''    End If
End
'Call Mexit_Click
End Sub

Private Sub MDIForm_Resize()
    timFl.Enabled = False
End Sub

Private Sub Mexit_Click()
Dim aa As Integer
    aa = MsgBox("是否真的要退出豪曼系统", vbInformation + vbYesNo)
    If aa = vbYes Then
        Call mod1.zhuLK '退出时取消注册
        End
End If
End Sub

Private Sub timGG_Timer()

End Sub

Private Sub timGGl_Timer()
Dim tt As String


On Error Resume Next
tt = "select top 1 gid,zz from ggl where " & mod1.DName & "=0 and datediff(second,rq,getdate())<5  order by gid desc"

End Sub

Private Sub timOnline_Timer()
Dim tt As String
On Error Resume Next

End Sub

Private Sub MFY_Click()

FydED.Show
FydED.ZOrder 0
End Sub

Private Sub timFl_Timer()
      Dim FlashInfo     As FLASHWINFO

            'Specifies   the   size   of   the   structure.
            FlashInfo.cbSize = Len(FlashInfo)
            'Specifies   the   flash   status
            FlashInfo.dwFlags = FLASHW_ALL Or FLASHW_TIMER
            'Specifies   the   rate,   in   milliseconds,   at   which   the   window   will   be   flashed.   If _
              dwTimeout   is   zero,   the   function   uses   the   default   cursor   blink   rate.
            FlashInfo.dwTimeout = 0
            'Handle   to   the   window   to   be   flashed.   The   window   can   be   either   opened   or   minimized.
            FlashInfo.hwnd = Me.hwnd
            'Specifies   the   number   of   times   to   flash   the   window.
            FlashInfo.uCount = 1
            FlashWindowEx FlashInfo
End Sub

Private Sub timXz_Timer()
On Error Resume Next
mod1.DQda = DateSerial(Year(mod1.DQda), Month(mod1.DQda), Day(mod1.DQda)) & " " & _
            TimeSerial(Hour(mod1.DQda), Minute(mod1.DQda), Second(mod1.DQda) + 1)

ztT.Panels(2).Text = mod1.DQda
End Sub

Private Sub ztT_PanelClick(ByVal Panel As MSComctlLib.Panel)
Dim oo As Integer
On Error Resume Next
If Panel.Index = 3 And Left(ztT.Panels(3).Text, 6) <> "您有未看信息" Then

        frmGGL.WindowState = 0
        frmGGL.ZOrder 0
ElseIf Panel.Index = 3 And Left(ztT.Panels(3).Text, 6) = "您有未看信息" Then
    MDI.timFl.Enabled = False
    Call frmOL.Tbound(Panel.Key)

    frmOL.Show
    frmOL.Left = frmZu.Left
    frmOL.Top = 0
    frmOL.ZOrder 0
    frmOL.Caption = "您正在和" & Panel.ToolTipText & "交谈"
    frmOL.img1.Picture = frmZu.ImageList2.ListImages(Panel.Tag).Picture
    frmOL.lbl1.Caption = Panel.ToolTipText

    frmOL.lbl1.ToolTipText = Panel.Key
    'frmOL.img2.Picture = frmZu.ImageList2.ListImages(frmZu.tb1.Buttons(frmZu.meIndex).Image).Picture
    'frmOL.img2.Picture = frmZu.ImageList2.ListImages(frmZu.NR(frmZu.meIndex)).Picture
    frmOL.img2.Picture = frmZu.NR(frmZu.meIndex).PictureNormal
    frmOL.lbl2.Caption = mod1.DName

    frmOL.txt2.Text = ""
    frmOL.txt2.SetFocus

'''''    For oo = 1 To frmZu.tb1.Buttons.Count
'''''        If frmZu.tb1.Buttons(oo).Key = Panel.Key Then
'''''            frmZu.tb1.Buttons(oo).Caption = frmZu.tb1.Buttons(oo).ToolTipText
'''''            Exit For
'''''        End If
'''''    Next
'''''    Panel.Text = ""
'''''    frmOL.txt1.SelStart = Len(frmOL.txt1.Text)
'''''    frmOL.txt1.SelLength = 0
    For oo = 1 To 50
        If frmZu.NR(oo).Tag = Panel.Key Then
            frmZu.NR(oo).Caption = frmZu.NR(oo).ToolTipText
            Exit For
        End If
    Next
    Panel.Text = ""
    frmOL.txt1.SelStart = Len(frmOL.txt1.Text)
    frmOL.txt1.SelLength = 0
End If
End Sub

