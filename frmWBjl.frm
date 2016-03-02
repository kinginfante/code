VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmWBjl 
   Caption         =   "维保记录表"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Frame frmRen 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   435
      Left            =   6840
      TabIndex        =   15
      Top             =   60
      Width           =   5805
      Begin VB.CommandButton cmdRe 
         Caption         =   "查 询"
         Height          =   285
         Left            =   4920
         TabIndex        =   20
         Top             =   90
         Width           =   795
      End
      Begin VB.ComboBox txtZ 
         Height          =   300
         Left            =   2640
         TabIndex        =   19
         Top             =   90
         Width           =   2235
      End
      Begin VB.ComboBox comLx 
         Height          =   300
         ItemData        =   "frmWBjl.frx":0000
         Left            =   990
         List            =   "frmWBjl.frx":000A
         TabIndex        =   17
         Text            =   "项目名称"
         Top             =   90
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "值:"
         Height          =   255
         Left            =   2310
         TabIndex        =   18
         Top             =   150
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "查询类别:"
         Height          =   225
         Left            =   90
         TabIndex        =   16
         Top             =   150
         Width           =   855
      End
   End
   Begin MSDataListLib.DataList comXmmc 
      Height          =   5940
      Left            =   330
      TabIndex        =   14
      Top             =   540
      Visible         =   0   'False
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   10478
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoNr 
      Height          =   330
      Left            =   2580
      Top             =   150
      Visible         =   0   'False
      Width           =   3465
      _ExtentX        =   6112
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
   Begin VB.Frame frmMod 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   30
      TabIndex        =   11
      Top             =   8640
      Width           =   13695
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         Height          =   345
         Left            =   1800
         TabIndex        =   13
         Top             =   150
         Width           =   1035
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加"
         Height          =   345
         Left            =   480
         TabIndex        =   12
         Top             =   150
         Width           =   1035
      End
   End
   Begin MSDataGridLib.DataGrid dtpNr 
      Bindings        =   "frmWBjl.frx":001E
      Height          =   8115
      Left            =   0
      TabIndex        =   7
      Top             =   510
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   14314
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "楷体_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "xmmc"
         Caption         =   "项目名称"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Y1"
         Caption         =   "1月份"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Y2"
         Caption         =   "2月份"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Y3"
         Caption         =   "3月份"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Y4"
         Caption         =   "4月份"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Y5"
         Caption         =   "5月份"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Y6"
         Caption         =   "6月份"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Y7"
         Caption         =   "7月份"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Y8"
         Caption         =   "8月份"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Y9"
         Caption         =   "9月份"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Y10"
         Caption         =   "10月份"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Y11"
         Caption         =   "11月份"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Y12"
         Caption         =   "12月份"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Button          =   -1  'True
            ColumnWidth     =   3704.882
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   900.284
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdMod 
      Height          =   375
      Left            =   13770
      Picture         =   "frmWBjl.frx":0032
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "修改"
      Top             =   8760
      Width           =   465
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   14250
      Picture         =   "frmWBjl.frx":033C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "提交"
      Top             =   8760
      Width           =   465
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14730
      Picture         =   "frmWBjl.frx":09A6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "返回"
      Top             =   8760
      Width           =   465
   End
   Begin VB.CommandButton cmdL 
      Caption         =   "<"
      Height          =   255
      Left            =   13620
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdR 
      Caption         =   ">"
      Height          =   225
      Left            =   14790
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblZNAME 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   1020
      TabIndex        =   10
      Top             =   150
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "组长:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Top             =   150
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "年份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   12990
      TabIndex        =   8
      Top             =   150
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "年"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   14520
      TabIndex        =   6
      Top             =   120
      Width           =   225
   End
   Begin VB.Label lblYear 
      Caption         =   "2006"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   14010
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmWBjl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoZu As ADODB.Recordset
Dim adoZZ As ADODB.Recordset


Private Sub cmdAdd_Click()
On Error Resume Next
adoNr.Recordset.AddNew "ywy", mod1.DName
adoNr.Recordset.Update "uid", mod1.DHid
adoNr.Recordset.Update "YY", Year(mod1.DQda)
Set dtpNr.DataSource = adoNr
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
frmZu.ZOrder 0
End Sub

Private Sub cmdDel_Click()
On Error Resume Next
adoNr.Recordset.Delete adAffectCurrent
End Sub

Private Sub cmdL_Click()
Dim tt As String
On Error Resume Next
lblYear.Caption = lblYear.Caption - 1
If frmRen.Visible = True Then
    If comLx.Text = "项目名称" Then
        tt = "select * from wbjl where xmmc like '%" & txtZ.Text & "%' and yy=" & lblYear.Caption & " order by xmmc,jid"
    ElseIf comLx.Text = "组长" Then
        tt = "select * from wbjl where ywy ='" & txtZ.Text & "' and yy=" & lblYear.Caption & " order by xmmc,jid"
    End If
Else
        tt = "select * from wbjl where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "' and yy=" & lblYear.Caption & " order by jid"
End If
frmWBjl.adoNr.Recordset.Close
If mod1.Zuf = True And mod1.DName <> "张寅" Then
    frmWBjl.adoNr.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    frmWBjl.cmdMod.Enabled = True
Else
    frmWBjl.adoNr.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    frmWBjl.cmdMod.Enabled = False
End If
Set frmWBjl.dtpNr.DataSource = frmWBjl.adoNr
End Sub

Private Sub cmdMod_Click()
frmMod.Visible = True
cmdSave.Enabled = True
End Sub

Private Sub cmdR_Click()
Dim tt As String
On Error Resume Next
lblYear.Caption = lblYear.Caption + 1
If frmRen.Visible = True Then
    If comLx.Text = "项目名称" Then
        tt = "select * from wbjl where xmmc like '%" & txtZ.Text & "%' and yy=" & lblYear.Caption & " order by xmmc,jid"
    ElseIf comLx.Text = "组长" Then
        tt = "select * from wbjl where ywy ='" & txtZ.Text & "' and yy=" & lblYear.Caption & " order by xmmc,jid"
    End If
Else
        tt = "select * from wbjl where ywy='" & mod1.DName & "' and uid='" & mod1.DHid & "' and yy=" & lblYear.Caption & " order by jid"
End If
frmWBjl.adoNr.Recordset.Close
If mod1.Zuf = True Then
    frmWBjl.adoNr.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    frmWBjl.cmdMod.Enabled = True
Else
    frmWBjl.adoNr.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    frmWBjl.cmdMod.Enabled = False
End If
Set frmWBjl.dtpNr.DataSource = frmWBjl.adoNr
End Sub

Private Sub DataGrid1_Click()

End Sub



Private Sub cmdRe_Click()
Dim tt As String
On Error Resume Next
If comLx.Text = "项目名称" Then
    tt = "select * from wbjl where xmmc like '%" & txtZ.Text & "%' and yy=" & Year(mod1.DQda) & " order by xmmc,jid"
    frmWBjl.adoNr.Recordset.Close
    frmWBjl.adoNr.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
ElseIf comLx.Text = "组长" Then
    lblZNAME.Caption = txtZ.Text
    tt = "select * from wbjl where ywy='" & txtZ.Text & "' and yy=" & Year(mod1.DQda) & " order by jid"
    If txtZ.Text = mod1.DName Then '如果选择为本人,则可以修改(针对王卫东)
        frmWBjl.adoNr.Recordset.Close
        frmWBjl.adoNr.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        cmdMod.Enabled = True
    Else
        frmWBjl.adoNr.Recordset.Close
        frmWBjl.adoNr.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        cmdMod.Enabled = False
    End If
End If
Set frmWBjl.dtpNr.DataSource = frmWBjl.adoNr
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
adoNr.Recordset.UpdateBatch
cmdSave.Enabled = False
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub comLx_Click()
Dim oo As Integer
On Error Resume Next
For oo = 15 To 0 Step -1
    txtZ.RemoveItem oo
Next
If comLx.Text = "组长" Then
    adoZZ.MoveFirst
    Do While Not adoZZ.EOF
        txtZ.AddItem adoZZ.Fields("username").Value
        adoZZ.MoveNext
    Loop
End If
End Sub


Private Sub comXmmc_DblClick()
adoNr.Recordset.Update "xmmc", comXmmc.Text
adoNr.Recordset.Update "htbh", comXmmc.BoundText
comXmmc.Visible = False
End Sub


Private Sub dtpNr_ButtonClick(ByVal ColIndex As Integer)

'comXmmc.Top = dtpNr.Row * dtpNr.RowHeight + dtpNr.Top + 500
comXmmc.Visible = True
End Sub

Private Sub dtpNr_Click()
comXmmc.Visible = False
End Sub


Private Sub Form_Click()
comXmmc.Visible = False
End Sub

Private Sub Form_Load()
Dim tt As String
On Error Resume Next
Me.Left = 0
Me.Top = 0
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
lblYear.Caption = Year(Date)
comXmmc.Visible = False
comXmmc.Left = 330
Set adoZu = New ADODB.Recordset

tt = "select xmmc,htbh from wbZname where zname='" & mod1.DName & "' and uid='" & mod1.DHid & "' and comid=" & mod1.comId & " order by xmmc"

adoZu.Close
'基础发布
'Select Case mod1.Lqy
'Case "上海"
'    adoZu.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    adoZu.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
adoZu.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
Set comXmmc.RowSource = adoZu
comXmmc.ListField = "xmmc"
comXmmc.BoundColumn = "htbh"

If mod1.comId = 0 Then
    tt = "select username,gzu from worker_gcz where zuf=1 order by gzu"
ElseIf mod1.comId = 1 Then
    tt = "select username,gzu from worker_gcz where zuf=3 and comid=1 order by gzu"
End If
Set adoZZ = New ADODB.Recordset
adoZZ.Close
'基础发布
'Select Case mod1.Lqy
'Case "上海"
'    adoZZ.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'Case "杭州"
'    adoZZ.Open tt, mod1.workHz, adOpenKeyset, adLockReadOnly, adCmdText
'End Select
adoZZ.Open tt, mod1.workFF, adOpenKeyset, adLockReadOnly, adCmdText
End Sub
