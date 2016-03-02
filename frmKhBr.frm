VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmKhBr 
   Caption         =   "项目资料查询"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15180
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   15180
   Begin VB.ComboBox comKhmc 
      Height          =   300
      Left            =   11460
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   900
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.TextBox txtKhdm 
      Height          =   345
      Left            =   12060
      TabIndex        =   19
      Text            =   "khdm"
      Top             =   2670
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.ComboBox comHyXz 
      Height          =   300
      ItemData        =   "frmKhBr.frx":0000
      Left            =   11460
      List            =   "frmKhBr.frx":0019
      TabIndex        =   17
      Top             =   1380
      Visible         =   0   'False
      Width           =   2505
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "新 建 项 目"
      Height          =   435
      Left            =   9990
      TabIndex        =   16
      Top             =   1890
      Width           =   3975
   End
   Begin VB.CommandButton cmdVall 
      Caption         =   "显示全部"
      Height          =   375
      Left            =   9960
      TabIndex        =   15
      Top             =   6990
      Width           =   1485
   End
   Begin VB.TextBox txtZ 
      Height          =   315
      Left            =   11370
      TabIndex        =   14
      Top             =   6120
      Width           =   2595
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "frmKhBr.frx":005B
      Left            =   11370
      List            =   "frmKhBr.frx":0065
      TabIndex        =   12
      Top             =   5460
      Width           =   2595
   End
   Begin TabDlg.SSTab tabCx 
      Height          =   9165
      Left            =   -30
      TabIndex        =   1
      Top             =   0
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   16166
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "按项目查询"
      TabPicture(0)   =   "frmKhBr.frx":007D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dtgKh"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frmPx"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "按联系人查询"
      TabPicture(1)   =   "frmKhBr.frx":0099
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPr(2)"
      Tab(1).Control(1)=   "cmdPr(1)"
      Tab(1).Control(2)=   "cmdPr(0)"
      Tab(1).Control(3)=   "dtgLx"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdPr 
         Caption         =   "项目名称"
         Height          =   375
         Index           =   2
         Left            =   -72420
         TabIndex        =   10
         Top             =   8760
         Width           =   5295
      End
      Begin VB.CommandButton cmdPr 
         Caption         =   "姓  别"
         Height          =   375
         Index           =   1
         Left            =   -73620
         TabIndex        =   9
         Top             =   8760
         Width           =   1185
      End
      Begin VB.CommandButton cmdPr 
         Caption         =   "客户姓名"
         Height          =   345
         Index           =   0
         Left            =   -74940
         TabIndex        =   8
         Top             =   8790
         Width           =   1305
      End
      Begin VB.Frame frmPx 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   465
         Left            =   0
         TabIndex        =   4
         Top             =   8730
         Width           =   9555
         Begin VB.CommandButton cmdPx 
            Caption         =   "行业性质"
            Height          =   405
            Index           =   2
            Left            =   7020
            TabIndex        =   7
            Top             =   30
            Width           =   1785
         End
         Begin VB.CommandButton cmdPx 
            Caption         =   "代号"
            Height          =   405
            Index           =   1
            Left            =   5850
            TabIndex        =   6
            Top             =   30
            Width           =   1155
         End
         Begin VB.CommandButton cmdPx 
            Caption         =   "业主资料"
            Height          =   405
            Index           =   0
            Left            =   60
            TabIndex        =   5
            Top             =   30
            Width           =   5775
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgKh 
         Height          =   8475
         Left            =   0
         TabIndex        =   2
         Top             =   330
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   14949
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLx 
         Height          =   8475
         Left            =   -74970
         TabIndex        =   3
         Top             =   300
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   14949
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "导航"
      Height          =   585
      Left            =   14460
      Picture         =   "frmKhBr.frx":00B5
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8580
      Width           =   675
   End
   Begin VB.Label Label3 
      Caption         =   "客户名称："
      Height          =   315
      Left            =   10380
      TabIndex        =   20
      Top             =   930
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label5 
      Caption         =   "客户行业性质："
      Height          =   285
      Left            =   10020
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label2 
      Caption         =   "值："
      Height          =   405
      Left            =   10740
      TabIndex        =   13
      Top             =   6150
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "查询条件："
      Height          =   225
      Left            =   10200
      TabIndex        =   11
      Top             =   5520
      Width           =   1065
   End
End
Attribute VB_Name = "frmKhBr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoKhBr As Object
Public adoRenBr As Object

Private Sub cmdBack_Click()

frmKhBr.Visible = False
frmZu.Enabled = True

End Sub

Private Sub cmdNew_Click()
Dim tt As String
Dim bt As String
On Error Resume Next
'''''''If cmdNew.ToolTipText = "" Then
'''''''    MsgBox "您的权限没有被完全设置,请速与马晓聪联系!"
'''''''    Exit Sub
'''''''End If


    Call mod1.xmQing
    Call mod1.khQing
    Call mod1.khRQing
    wbDN.Show

    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.CC
    mod1.cmd.CommandText = "XMjia"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@Uid") = mod1.DHid

    mod1.cmd.Execute
    wbDN.lblXid.Caption = mod1.cmd.Parameters("@xid").Value

    wbDN.lblYwy.Caption = mod1.DName
    wbDN.lblUid.Caption = mod1.DHid
    wbDN.LCRen = mod1.DName
    wbDN.LCUid = mod1.DHid
    wbDN.Lc = 1
    wbDN.lblLcRen.Caption = mod1.DName
    wbDN.lblLcUid.Caption = mod1.DHid
    wbDN.lblXmPd.Caption = 0
    Set cmd = Nothing
'    wbDN.lblXywy.Caption = mod1.DName
'    wbDN.lblXuid.Caption = mod1.DHid
    Call mod1.KhJQing
 

    wbDN.tabKh.Tab = 0
    wbDN.tabKh.TabEnabled(1) = False
    wbDN.tabKh.Enabled = False
    wbDN.optYz.Value = False
    wbDN.optWy.Value = False
    wbDN.optQt.Value = False
    wbDN.frmGL.Visible = False
    wbDN.frmJz.Visible = True
    
    wbDN.cmdNew.Enabled = False
    wbDN.cmdRdel.Enabled = False

    comHyxz.Text = ""
    txtKhdm.Text = ""

    wbDN.khAdd = True '为新建项目
    If frmKhBr.Visible = True Then
        frmKhBr.Enabled = False
    End If
    Set wbDN.txtKhmc.RowSource = Nothing
    Call mod1.XmKhUnLocked
    wbDN.dtgP.Visible = False
End Sub

Private Sub cmdPr_Click(Index As Integer)
On Error Resume Next
Static aa As Boolean
dtgLx.Col = Index + 1
If aa = True Then
    dtgLx.Sort = 1
    aa = False
Else
    dtgLx.Sort = 2
    aa = True
End If
End Sub

Private Sub cmdPx_Click(Index As Integer)
Static aa As Boolean
dtgKh.Col = Index + 1
If aa = True Then
    dtgKh.Sort = 1
    aa = False
Else
    dtgKh.Sort = 2
    aa = True
End If

End Sub

Private Sub Command1_Click()

End Sub








Private Sub cmdVall_Click()
Dim tt As String
On Error Resume Next
    tt = "VkhrNew('" & mod1.DName & "')"
    frmKhBr.adoRenBr.Close
    frmKhBr.adoRenBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    Set frmKhBr.dtgLx.DataSource = frmKhBr.adoRenBr
    tabCx.Tab = 1
End Sub


Private Sub dtgKH_DblClick()
Dim tt As String
On Error Resume Next
Dim Kid As Long
Dim xid As Long

    dtgKh.Col = 2
    xid = Val(dtgKh.Text)
    dtgKh.Col = 5
    Kid = Val(dtgKh.Text)
    dtgKh.Col = 2
    If dtgKh.Text = "" Then Exit Sub
    frmWait.Show
    frmWait.ZOrder 0
    
    frmWait.Refresh
    frmWait.faWait.Play
    


    'wbDN.WindowState = 2

    'mod1.Kd = frmKhBrow.adoKh.Recordset.Fields("khdh").Value
    If mod1.DKZ(xid, 7) = True Then
        MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
        Exit Sub
    End If
    
    frmKhBr.Enabled = False
    wbDN.Visible = False
    Me.MousePointer = 11
    '记录打开日志
    Call mod1.zhuDa(7, dtgKh.Text)
    Call mod1.xmQing
    Call mod1.khQing
    Call mod1.xmBound(xid)
    wbDN.lblKid.Caption = wbDN.lblYZ.Tag
    Call mod1.khBound(wbDN.lblYZ.Tag, "yz")
    If Val(wbDN.lblXmPd.Caption) < 60 Then
        wbDN.frmJE.Visible = True
    End If
    wbDN.Left = 0
    wbDN.Top = 0
    wbDN.cmdMod.Enabled = True
    wbDN.cmdSave.Enabled = False
    Me.MousePointer = 0
    wbDN.tabKh.Tab = 0
    'wbDN.cmdRadd.Enabled = True
    If wbDN.txtKhmc.Text = "" Then
        wbDN.tabKh.TabEnabled(1) = False
    Else
        wbDN.tabKh.TabEnabled(1) = True
    End If
    wbDN.tabKh.TabEnabled(0) = True
    wbDN.cmdSave.Enabled = True
    'wbDN.cmdSaveA.Enabled = True
    
    
    

    wbDN.modFi = False

    Me.MousePointer = 0
    wbDN.cmdSave.Enabled = False
    wbDN.tabKh.Enabled = True
    If wbDN.lblYwy.Caption = mod1.DName Or wbDN.lblXywy.Caption = mod1.DName Then
        wbDN.cmdMod.Enabled = True
    Else
        wbDN.cmdMod.Enabled = False
    End If
    wbDN.khAdd = False
    '打开项目后,默认的打开客户为项目资料
    wbDN.optYz.Value = True
    wbDN.frmGL.Visible = False
    wbDN.frmJz.Visible = True
    frmWait.Visible = False
    wbDN.Visible = True
    wbDN.cmdMod.Enabled = True
    
    '更新动态签字按钮的初始设置
        For oo = 1 To 10
           wbDN.lblQM(oo).Left = wbDN.lblQM(oo - 1).Left + 1100
           wbDN.cmdQm(oo).Left = wbDN.cmdQm(oo - 1).Left + 1100
           wbDN.lblTm(oo).Left = wbDN.lblTm(oo - 1).Left + 1100
           mod1.HTP.MoveNext
        Next
End Sub


Private Sub dtgLx_DblClick()
'Dim tt As String
'On Error Resume Next
'Dim Kid As Double
'Dim xid As Double
'
'    dtgLx.Col = 4
'    Kid = dtgLx.Text
'    dtgLx.Col = 5
'    xid = dtgLx.Text
'    dtgLx.Col = 2
'    If dtgLx.Text = "" Then Exit Sub
'    frmWait.Show
'    frmWait.ZOrder 0
'    frmKhBr.Enabled = False
'    wbDN.Show
'    'wbDN.WindowState = 2
'    Me.MousePointer = 11
'    'mod1.Kd = frmKhBrow.adoKh.Recordset.Fields("khdh").Value
'
'    If mod1.DKZ(Kid, 6) = True Then
'        MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
'        Exit Sub
'    End If
'
'    '记录打开日志
'    Call mod1.zhuDa(3, dtgLx.Text)
'
'    Call mod1.khQing
'    Call mod1.khBound(Kid, xid)
'
'    wbDN.Left = 0
'    wbDN.Top = 0
'    wbDN.cmdMod.Enabled = True
'    wbDN.cmdSave.Enabled = False
'    Me.MousePointer = 0
'    wbDN.tabKh.Tab = 0
'    wbDN.cmdRadd.Enabled = True
'    wbDN.tabKh.TabEnabled(2) = True
'    wbDN.tabKh.TabEnabled(0) = True
'    wbDN.cmdSave.Enabled = True
'    'wbDN.cmdSaveA.Enabled = True
'
'    wbDN.cmdAddA.Enabled = True
'    wbDN.cmdDelA.Enabled = True
'    wbDN.cmdAddB.Enabled = True
'    wbDN.cmdDelB.Enabled = True
'    wbDN.cmdAddC.Enabled = True
'    wbDN.cmdDelC.Enabled = True
'    wbDN.cmdAddD.Enabled = True
'    wbDN.cmdDelD.Enabled = True
'    wbDN.cmdAddE.Enabled = True
'    wbDN.cmdDelE.Enabled = True
'    wbDN.cmdAddF.Enabled = True
'    wbDN.cmdDelF.Enabled = True
'    wbDN.cmdAddG.Enabled = True
'    wbDN.cmdDelG.Enabled = True
'
'    wbDN.dtgA.AllowUpdate = True
'    wbDN.dtgB.AllowUpdate = True
'    wbDN.dtGC.AllowUpdate = True
'    wbDN.dtgD.AllowUpdate = True
'    wbDN.dtgE.AllowUpdate = True
'    wbDN.dtgF.AllowUpdate = True
'    wbDN.dtgG.AllowUpdate = True
'    wbDN.modFi = False
'    frmWait.Visible = False
'    wbDN.khAdd = False
End Sub


Private Sub Form_Load()
frmKhBr.Height = mod1.FHeight
frmKhBr.Width = mod1.FWidth
dtgKh.ColWidth(0) = 300
dtgKh.ColWidth(1) = 4600
'dtgKH.ColWidth(4) = 700
'dtgKh.ColWidth(5) = 0
dtgKh.ColWidth(6) = 0
dtgKh.ColWidth(7) = 0
dtgKh.ColWidth(8) = 0
dtgKh.ColWidth(9) = 0
dtgKh.ColWidth(10) = 0
dtgKh.ColWidth(11) = 0

dtgLx.ColWidth(0) = 300
dtgLx.ColWidth(3) = 5500
dtgLx.ColWidth(4) = 0
dtgLx.ColWidth(5) = 0
'Call ResizeInit(Me) '在程序装入时必须加入
End Sub

Private Sub Form_Resize()
'cmdBack.Left = frmKhBr.Width - cmdBack.Width - 500
'cmdBack.Top = frmKhBr.Height - cmdBack.Height - 700
'dtgKh.Height = frmKhBr.Height - 1300
'frmPx.Top = dtgKh.Height + 100
End Sub

Private Sub opTa_Click()
dtgKh.Col = 1
dtgKh.Sort = 1
End Sub

Private Sub optB_Click()
dtgKh.Col = 2
dtgKh.Sort = 2
End Sub

Private Sub optC_Click()
dtgKh.Col = 3
dtgKh.Sort = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
frmKhBr.Visible = False
frmZu.Enabled = True
Cancel = True
End If
End Sub

Private Sub txtZ_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tt As String
On Error Resume Next
If KeyCode = 13 Then
    Select Case comLx.Text
    Case "项目名称"
        tt = "khNewV_xmmc('" & mod1.DName & "','" & txtZ.Text & "')"
        frmKhBr.adoKhBr.Close
        frmKhBr.adoKhBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        Set frmKhBr.dtgKh.DataSource = frmKhBr.adoKhBr
        tabCx.Tab = 0
    Case "客户姓名"
        tt = "khNewV_man('" & mod1.DName & "','" & txtZ.Text & "')"
        frmKhBr.adoRenBr.Close
        frmKhBr.adoRenBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        Set frmKhBr.dtgLx.DataSource = frmKhBr.adoRenBr
        tabCx.Tab = 1
    End Select

End If
End Sub
