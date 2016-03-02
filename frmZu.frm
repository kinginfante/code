VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmZu 
   Caption         =   "业务导航图"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9345
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   9345
   Begin TabDlg.SSTab tabZu 
      Height          =   4545
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   8017
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "维保业务"
      TabPicture(0)   =   "frmZu.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdBu(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "产品业务"
      TabPicture(1)   =   "frmZu.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBb(0)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "服务"
      TabPicture(2)   =   "frmZu.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdBc(0)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "行政"
      TabPicture(3)   =   "frmZu.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdBd(0)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "工程"
      TabPicture(4)   =   "frmZu.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdBe(0)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "财务"
      TabPicture(5)   =   "frmZu.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdBf(0)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "管理"
      TabPicture(6)   =   "frmZu.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "cmdBg(0)"
      Tab(6).ControlCount=   1
      Begin VB.CommandButton cmdBf 
         Height          =   1005
         Index           =   0
         Left            =   -73830
         TabIndex        =   8
         Top             =   1380
         Width           =   1155
      End
      Begin VB.CommandButton cmdBe 
         Height          =   1005
         Index           =   0
         Left            =   -73860
         TabIndex        =   7
         Top             =   1230
         Width           =   1155
      End
      Begin VB.CommandButton cmdBd 
         Height          =   1005
         Index           =   0
         Left            =   -73950
         TabIndex        =   6
         Top             =   1230
         Width           =   1155
      End
      Begin VB.CommandButton cmdBc 
         Height          =   1005
         Index           =   0
         Left            =   -73980
         TabIndex        =   5
         Top             =   1170
         Width           =   1155
      End
      Begin VB.CommandButton cmdBb 
         Height          =   1005
         Index           =   0
         Left            =   -74190
         TabIndex        =   4
         Top             =   1380
         Width           =   1155
      End
      Begin VB.CommandButton cmdBu 
         Height          =   1005
         Index           =   0
         Left            =   750
         TabIndex        =   3
         Top             =   1440
         Width           =   1155
      End
      Begin VB.CommandButton cmdBg 
         Height          =   1005
         Index           =   0
         Left            =   -74340
         TabIndex        =   2
         Top             =   990
         Width           =   1155
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3780
      Top             =   6570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   41
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZu.frx":00C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZu.frx":2E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZu.frx":5FBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZu.frx":8E65
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZu.frx":BBDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZu.frx":E8AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBa 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   4605
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   1296
      ButtonWidth     =   1244
      ButtonHeight    =   1244
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "公告栏"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "报销"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "事务"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "信息"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "企业文化"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Width           =   4500
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmZu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cmdBu_Click(Index As Integer)
Dim tt As String
On Error Resume Next
'MsgBox cmdBu(Index).Caption
Select Case Index
Case 0 '业主资料
    frmKhBr.Show
    frmKhBr.Enabled = True
    frmKhBr.ZOrder 0
    Set frmKhBr.adoKhBr = New ADODB.Recordset
    tt = "vkhNew('" & mod1.DName & "','" & mod1.DHid & "')"
    frmKhBr.adoKhBr.Close
    frmKhBr.adoKhBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    Set frmKhBr.dtgKh.DataSource = frmKhBr.adoKhBr
End Select
End Sub

Private Sub Form_Load()
frmZu.Width = 9465
frmZu.Height = 5910
cmdBu(0).Left = 1110
cmdBu(0).Top = 870

cmdBb(0).Left = 1110
cmdBb(0).Top = 870
'cmdBc(0).Left = 1110
'cmdBc(0).Top = 870
cmdBd(0).Left = 1110
cmdBd(0).Top = 870
cmdBe(0).Left = 1110
cmdBe(0).Top = 870
cmdBf(0).Left = 1110
cmdBf(0).Top = 870
cmdBg(0).Left = 1110
cmdBg(0).Top = 870

'Call ResizeInit(Me) '在程序装入时必须加入
End Sub

Private Sub TabStrip1_Change()

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub Form_Resize()
'frmZu.WindowState = 2
'Call mod1.ResizeForm(Me) '确保窗体改变时控件随之改变
tabZu.Width = frmZu.Width
End Sub

Private Sub tabZu_DblClick()

End Sub

Private Sub TBa_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tt As String
Dim oo As Integer
Dim PK As String
On Error Resume Next
'MsgBox Button.Index
Select Case Button.Index
    Case 2 '公告栏
    
        frmGGL.Show
        frmGGL.cmdSave.Enabled = False
        frmGGL.ZOrder 0

    Case 3 '报销
        mod1.BTZ = 23
        frmBxBrow.WindowState = 0
        frmBxBrow.Show
        frmBxBrow.WindowState = 2
        Set frmBxBrow.AdoBxBro = New ADODB.Recordset
        tt = "FydV('" & mod1.AdoDlYwy.Fields("uid").Value & "','" & AdoDlYwy.Fields("ywy").Value & "')"
        frmBxBrow.AdoBxBro.Close
        frmBxBrow.AdoBxBro.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adcmdstoreproc
        Set frmBxBrow.mga.DataSource = frmBxBrow.AdoBxBro
        'PK = "<起 始 期  |<截 至 期  |>  金 额 |^ 报 销 单 编 号|> 签收日期 "
        'PK = "^  日期范围|^  日期范围|>  金 额 |^ 报 销 单 编 号|> 签收日期 "
         frmBxBrow.mga.ColWidth(0) = 500
        'frmBxBrow.mga.FormatString = PK
        frmBxBrow.mga.MergeRow(0) = True
        frmBxBrow.mga.MergeCells = flexMergeRestrictAll
        frmBxBrow.optMe.Value = True
        'frmBxBrow.BorderStyle = 1
        'frmBxBrow.MGa.FixedCols = 1
'        frmBxBrow.MaxButton = False
'        frmBxBrow.MinButton = False

        '生成报销单按钮
        mod1.FydBut.MoveFirst
        For oo = 10 To 1 Step -1
            Unload frmBxBrow.cmdFyd(oo)
        Next
        frmBxBrow.cmdFyd(0).Caption = Trim(mod1.FydBut.Fields("lb").Value) & "报销单"
        frmBxBrow.cmdFyd(0).Tag = mod1.FydBut.Fields("mid").Value
        frmBxBrow.cmdFyd(0).ToolTipText = Trim(mod1.FydBut.Fields("Bz").Value) & ",流程的总数为:" & _
        mod1.FydBut.Fields("Lcou").Value
        mod1.FydBut.MoveNext
        'For oo = 1 To mod1.FydBut.RecordCount - 1
        oo = 1
        Do While Not mod1.FydBut.EOF
            '如果有重复的名称(主要为<=500)的钱),则不显示按钮
            If Not (mod1.FydBut.Fields("mid").Value = 47 Or _
            mod1.FydBut.Fields("mid").Value = 68 Or mod1.FydBut.Fields("mid").Value = 81 Or _
            mod1.FydBut.Fields("mid").Value = 141) Then
            Load frmBxBrow.cmdFyd(oo)
            frmBxBrow.cmdFyd(oo).Caption = Trim(mod1.FydBut.Fields("lb").Value) & "报销单"
            frmBxBrow.cmdFyd(oo).Tag = mod1.FydBut.Fields("mid").Value
            frmBxBrow.cmdFyd(oo).ToolTipText = Trim(mod1.FydBut.Fields("Bz").Value) & ",流程的总数为:" & _
            mod1.FydBut.Fields("Lcou").Value
               frmBxBrow.cmdFyd(oo).Top = frmBxBrow.cmdFyd(oo - 1).Top + 1500

                frmBxBrow.cmdFyd(oo).Visible = True
            oo = oo + 1
            End If
            mod1.FydBut.MoveNext
        'Next
        Loop
    Case 4 '当前事务
        Set Dialog.AdoDi = New ADODB.Recordset
        tt = "EnventOpen('" & mod1.DName & "','" & mod1.DHid & "')"
        Dialog.AdoDi.Close
        Dialog.AdoDi.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        Set Dialog.dtgDi.DataSource = Dialog.AdoDi
        Dialog.Show
        Dialog.ZOrder 0
        frmZu.Enabled = False
    Case 14
        MDI.Cq = True
        Unload MDI
        mod1.FiR = False
        Form1.Show
        Form1.Fa1.GotoFrame (160)
        Call mod1.zhuLK '退出时取消注册
End Select

End Sub
