VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmGzJ 
   Caption         =   "工作计划"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除"
      Enabled         =   0   'False
      Height          =   585
      Left            =   13920
      Picture         =   "frmGzJ.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8550
      Width           =   645
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14580
      Picture         =   "frmGzJ.frx":018A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8550
      Width           =   585
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "提交"
      Height          =   585
      Left            =   13230
      Picture         =   "frmGzJ.frx":028C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8550
      Width           =   675
   End
   Begin VB.Frame frmMod 
      BorderStyle     =   0  'None
      Height          =   7305
      Left            =   90
      TabIndex        =   3
      Top             =   780
      Width           =   15105
      Begin MSAdodcLib.Adodc adoFy 
         Height          =   405
         Left            =   7920
         Top             =   5520
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   714
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
      Begin MSAdodcLib.Adodc adoXmgz 
         Height          =   405
         Left            =   4590
         Top             =   5070
         Visible         =   0   'False
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   714
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
      Begin VB.TextBox txtXmFy 
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   150
         TabIndex        =   13
         Top             =   5640
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txtjzDC 
         Height          =   825
         Left            =   1290
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   2310
         Width           =   13725
      End
      Begin VB.TextBox txtXm 
         Height          =   795
         Left            =   1290
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1440
         Width           =   13725
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   705
         Left            =   1290
         TabIndex        =   6
         Top             =   0
         Width           =   8745
         Begin VB.OptionButton optA 
            Caption         =   "很困难(0)"
            Height          =   255
            Left            =   7110
            TabIndex        =   10
            Top             =   150
            Width           =   1185
         End
         Begin VB.OptionButton optB 
            Caption         =   "有难度(30)"
            Height          =   315
            Left            =   4900
            TabIndex        =   9
            Top             =   120
            Width           =   1245
         End
         Begin VB.OptionButton optC 
            Caption         =   "有可能(60)"
            Height          =   315
            Left            =   2690
            TabIndex        =   8
            Top             =   120
            Width           =   1245
         End
         Begin VB.OptionButton optD 
            Caption         =   "有把握(90)"
            Height          =   285
            Left            =   480
            TabIndex        =   7
            Top             =   150
            Width           =   1245
         End
      End
      Begin VB.TextBox txtXDBZ 
         Height          =   2415
         Left            =   1290
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   4860
         Width           =   13725
      End
      Begin VB.TextBox txtBfMd 
         Height          =   1575
         Left            =   1290
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3210
         Width           =   13725
      End
      Begin VB.Label Label11 
         Caption         =   "竞争对手："
         Height          =   285
         Left            =   30
         TabIndex        =   18
         Top             =   2370
         Width           =   1845
      End
      Begin VB.Label Label10 
         Caption         =   "项目描述："
         Height          =   285
         Left            =   30
         TabIndex        =   17
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "行动步骤："
         Height          =   675
         Left            =   30
         TabIndex        =   16
         Top             =   4920
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "拜访目的："
         Height          =   315
         Left            =   30
         TabIndex        =   15
         Top             =   3240
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "客户平台"
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   150
         Width           =   795
      End
   End
   Begin VB.TextBox txtzgPd 
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   8130
      Width           =   6285
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "修改"
      Height          =   585
      Left            =   12540
      Picture         =   "frmGzJ.frx":08F6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8550
      Width           =   675
   End
   Begin VB.CommandButton cmdKhzl 
      Caption         =   "客户资料"
      Height          =   285
      Left            =   8370
      TabIndex        =   0
      Top             =   450
      Width           =   1815
   End
   Begin VB.Label lblYwy 
      Caption         =   "Label20"
      Height          =   285
      Left            =   5850
      TabIndex        =   36
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label19 
      Caption         =   "业 务 员："
      Height          =   255
      Left            =   4890
      TabIndex        =   35
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblKhmc 
      DataField       =   "khQc"
      Height          =   285
      Left            =   1320
      TabIndex        =   34
      Top             =   90
      Width           =   3465
   End
   Begin VB.Label lblDm 
      Height          =   225
      Left            =   5850
      TabIndex        =   33
      Top             =   120
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label3 
      Caption         =   "客户代码："
      Height          =   225
      Left            =   4890
      TabIndex        =   32
      Top             =   120
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "项目名称："
      Height          =   345
      Left            =   120
      TabIndex        =   31
      Top             =   90
      Width           =   945
   End
   Begin VB.Label Label7 
      Caption         =   "日    期："
      Height          =   285
      Left            =   7410
      TabIndex        =   30
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblRq 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dddddd aaaa"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   8370
      TabIndex        =   29
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label lblWeek 
      Caption         =   "五"
      Height          =   225
      Left            =   9930
      TabIndex        =   28
      Top             =   120
      Width           =   225
   End
   Begin VB.Label Label9 
      Caption         =   "星期"
      Height          =   225
      Left            =   9540
      TabIndex        =   27
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   "地    址："
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   510
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblAdr 
      Height          =   255
      Left            =   1320
      TabIndex        =   25
      Top             =   510
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.Label lblZGQZ 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   150
      TabIndex        =   24
      Top             =   8610
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "主管评定："
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   8280
      Width           =   1125
   End
   Begin VB.Label Label16 
      Caption         =   "Label16"
      DataField       =   "UserId"
      DataSource      =   "adoXmgz"
      Height          =   195
      Left            =   6510
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "frmGzJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
Dim ii As Integer
Dim tt As String
On Error Resume Next
If cmdSave.Enabled = True Then
ii = MsgBox("退出将不保存数据！", vbYesNo + vbInformation, "当心！")
    If ii = vbNo Then Exit Sub
End If
frmZu.Enabled = True

If txtXDBZ.Text = "" And modXmGz.Ti = True Then
    tt = "delete from xmgz where gid=" & modXmGz.Gid
    adoXmgz.Recordset.Close
    adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adCmdText
'    adoXmgz.Recordset.Delete adAffectCurrent
'    adoXmgz.Recordset.UpdateBatch
    
End If
frmGzJ.Visible = False
End Sub

Private Sub cmdMod_Click()
frmGzJ.frmMod.Enabled = True
frmGzJ.cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
Dim tt As String
If txtXDBZ.Text = "" Then
    MsgBox "请输入行动步骤！"
    Exit Sub
End If
Call modXmGz.jiAdd
cmdSave.Enabled = False

'更新工作报告表
'tt = "Select * from xmgz where ywy like '%" & frmGzBG.comYwy.Text & "%' and aTime>='" & modXmGz.Fr & _
'"' and aTime <='" & modXmGz.Lr & "' and lb=0 order by aTime"
'frmGzBG.adoJi.Close
'frmGzBG.adoJi.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
frmGzBG.adoJi.Requery
Set frmGzBG.dtgJi.DataSource = frmGzBG.adoJi

End Sub


Private Sub Form_Load()
frmGzJ.Width = mod1.Fwidth
frmGzJ.Height = mod1.FHeight
frmGzJ.Left = 0
frmGzJ.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim tt As String
Dim ii As Integer
On Error Resume Next
If MDI.Cq = False Then
If cmdSave.Enabled = True Then
ii = MsgBox("退出将不保存数据！", vbYesNo + vbInformation, "当心！")
    If ii = vbNo Then Exit Sub
End If

frmZu.Enabled = True

If txtXDBZ.Text = "" And modXmGz.Ti = True Then
    tt = "delete from xmgz where gid=" & modXmGz.Gid
    adoXmgz.Recordset.Close
    adoXmgz.Recordset.Open tt, mod1.workKK, adOpenKeyset, adCmdText
'    adoXmgz.Recordset.Delete adAffectCurrent
'    adoXmgz.Recordset.UpdateBatch
    
End If
End If
End Sub

Private Sub txtzgPd_Click()
    If mod1.KhK >= 1 Then
    lblZGQZ.Caption = mod1.DName
    End If
End Sub


