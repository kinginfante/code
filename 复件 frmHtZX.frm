VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmHtZX 
   Caption         =   "合同执行列表"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Frame Frame3 
      Caption         =   "派工单"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4335
      Left            =   9030
      TabIndex        =   11
      Top             =   4080
      Width           =   6165
   End
   Begin VB.Frame Frame1 
      Caption         =   "配料单"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4065
      Left            =   9060
      TabIndex        =   10
      Top             =   30
      Width           =   6135
      Begin MSDataGridLib.DataGrid dtgPld 
         Height          =   3405
         Left            =   0
         TabIndex        =   17
         Top             =   210
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   6006
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "开单日期"
            Caption         =   "开单日期"
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
            DataField       =   "编号"
            Caption         =   "编号"
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
            DataField       =   "项目名称"
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
         BeginProperty Column03 
            DataField       =   "GuID"
            Caption         =   "GuID"
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
            DataField       =   "htBh"
            Caption         =   "htBh"
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
            DataField       =   "ZT"
            Caption         =   "流程"
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
            DataField       =   "lc"
            Caption         =   "lc"
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
            DataField       =   "ywy"
            Caption         =   "ywy"
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
            DataField       =   "uid"
            Caption         =   "uid"
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
               Object.Visible         =   -1  'True
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   2505.26
            EndProperty
            BeginProperty Column03 
               Object.Visible         =   0   'False
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column07 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "所 有"
         Height          =   315
         Left            =   2850
         TabIndex        =   16
         Top             =   3690
         Width           =   1005
      End
      Begin VB.CommandButton cmdPldOpen 
         Caption         =   "打  开"
         Height          =   315
         Left            =   5010
         TabIndex        =   15
         Top             =   3690
         Width           =   1065
      End
      Begin VB.CommandButton cmdNP 
         Caption         =   "新  建"
         Height          =   315
         Left            =   3840
         TabIndex        =   14
         Top             =   3690
         Width           =   1155
      End
      Begin VB.Label lblHtbh 
         Height          =   225
         Left            =   930
         TabIndex        =   13
         Top             =   3750
         Width           =   2265
      End
      Begin VB.Label Label3 
         Caption         =   "合同编号"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   3750
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdHtOpen 
      Caption         =   "打开"
      Height          =   345
      Left            =   7440
      TabIndex        =   9
      Top             =   8730
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "所有执行合同"
      Height          =   375
      Left            =   5850
      TabIndex        =   8
      Top             =   8730
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      Caption         =   "条件查询"
      Height          =   615
      Left            =   30
      TabIndex        =   2
      Top             =   8520
      Width           =   5745
      Begin VB.CommandButton cmdRef1 
         Caption         =   "查  询"
         Height          =   285
         Left            =   4590
         TabIndex        =   5
         Top             =   270
         Width           =   1035
      End
      Begin VB.ComboBox comXZ 
         Height          =   300
         ItemData        =   "frmHtZX.frx":0000
         Left            =   810
         List            =   "frmHtZX.frx":000D
         TabIndex        =   4
         Text            =   "合同金额"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtYc 
         Height          =   285
         Left            =   2820
         TabIndex        =   3
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "条件"
         Height          =   255
         Left            =   300
         TabIndex        =   7
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "值"
         Height          =   255
         Left            =   2610
         TabIndex        =   6
         Top             =   270
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "导航"
      Height          =   585
      Left            =   14520
      Picture         =   "frmHtZX.frx":002F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8580
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8445
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   14896
      _Version        =   393216
      BackColor       =   -2147483634
      BackColorBkg    =   -2147483636
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmHtZX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoBr As ADODB.Recordset
Public adoPld As ADODB.Recordset
Dim NewF As Boolean

Private Sub cmdAll_Click()
Dim tt As String
Dim Ywy As String
Dim Uid As String
On Error Resume Next

Ywy = mod1.DName
Uid = mod1.DHid

tt = "select * from PldView where ywy='" & Ywy & "' and uid='" & Uid & "' order by lc"
frmHtZX.adoPld.Close
frmHtZX.adoPld.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmHtZX.dtgPld.DataSource = frmHtZX.adoPld
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
End Sub

Private Sub cmdHtOpen_Click()

Dim tt As String
Dim xZ As String

Dim Hid As Long
'Dim Lid As String
On Error Resume Next
mod1.BTZ = 6
dtgBr.Col = 3
xZ = dtgBr.Text
dtgBr.Col = 6
Hid = dtgBr.Text
dtgBr.Col = 7
NewF = dtgBr.Text
'Lid = Str(Lid)
If mod1.DKZ(Hid, 1) = True Then
        MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
        Exit Sub
End If

frmWait.Visible = True
frmWait.ZOrder 0
frmWait.Refresh
'htBrow.MousePointer = 11
htBrow.Enabled = False
'mod1.MPld = False '初始化,不生成配料单
If NewF = False Then
    If xZ = "C. 维保合同" Or xZ = "D. 维修合同" Then
    'mod1.comJZ = False
    wbHTP.Visible = False
    Call modHt.wbQing
    
    
    tt = "Select * from htping where hid=" & Hid
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Call modHt.wbBound
    
    
    '打开材料表
    tt = "Select * from htSale where htbh='" & wbHTP.txtHtbh.Text & "'"
    wbMx.adoRGF.Recordset.Close
    wbMx.adoRGF.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set wbMx.dtgSale.DataSource = wbMx.adoRGF
    wbMx.lblChg.Caption = wbHTP.txtClcb1.Text
    
    '打开应收款表
    tt = "Select * from htping1 where htBh='" & wbHTP.txtHtbh.Text & "' order by rq"
    frmFuK.adoHpt.Recordset.Close
    frmFuK.adoHpt.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set wbMx.dtgFk.DataSource = frmFuK.adoHpt
    
    '打开佣金表
    tt = "Select * from Yongjin where htBh='" & wbHTP.txtHtbh.Text & "' order by yId"
    frmYj.adoYj.Recordset.Close
    frmYj.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set frmYj.dtgYj.DataSource = frmYj.adoYj
    
    ''打开出工信息表(如果为评审阶段则不显示）
    'If wbHTP.optZ.Value = True Or wbHTP.optW.Value = True Then
    '    tt = "Select max(gzb.rq),max(gzb.wxWorker),sum(workXX.wTime),max(bhid)" & _
    '    "max(htbh) from gzb cross join workXX where gzb.bhid=workXX.bhid and gzb.htBh='" & _
    '    wbHTP.txtHtbh.Text & "' group by gzb.bhid"
    '    form2Htp.adoGzb.Recordset.Close
    '    form2Htp.adoGzb.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '    Set wbMx.dtgGzb.DataSource = form2Htp.adoGzb
    'End If
    wbHTP.Visible = True
    
    wbHTP.txtYj1.Visible = False
    wbHTP.txtYj2.Visible = False
    wbHTP.txtLr1.Visible = False
    wbHTP.txtLr2.Visible = False
    wbHTP.lblTcBe.Visible = False
    wbHTP.txtTcBe.Visible = False
    wbHTP.UpDa.Visible = False
    wbHTP.lblYj.Visible = False
    wbHTP.lblLr.Visible = False
    wbHTP.lblTC.Visible = False
    Exit Sub
    End If
    
    
    
    
    
    
    
    
    
    
    '购销合同
    
    form2Htp.Visible = True
    mod1.workTt = ""
    mod1.workTt = "Select * from htPing where hid=" & Hid
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open mod1.workTt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    form2Htp.lblHtxz.Caption = ""
    
    Call modHt.htQing
    Call modHt.htBound '绑定合同评审单字段
    

    
    
    '打开收款表
    
    
    tt = "Select * from htPing1 where htBh='" & form2Htp.txtHtbh.Text & "' order by rq"
    frmFuK.adoHpt.Recordset.Close
    frmFuK.adoHpt.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    
    
    Set frmFuK.dtgFk.DataSource = frmFuK.adoHpt
    
    'ft = "Select * from yiFk Where htBh='" & frmFuK.adoHpt.Recordset.Fields("htBh").Value & _
    '"' and yingRQ='" & frmFuK.adoHpt.Recordset.Fields("rq").Value & "' order by yiRq"
    'frmFuK.adoYf.Recordset.Close
    'frmFuK.adoYf.Recordset.Open ft, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'Set frmFuK.dtgYf.DataSource = frmFuK.adoYf
    
    '打开产品表
    tt = ""
    tt = "Select * from htSale Where htBh='" & form2Htp.txtHtbh.Text & "'"
    form2Htp.adoSale.Recordset.Close
    form2Htp.adoSale.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set form2Htp.dtgSale.DataSource = form2Htp.adoSale
    Set form2Htp.dtgYj.DataSource = form2Htp.adoSale
    Set form2Htp.dtgZj.DataSource = form2Htp.adoSale
    
    ''打开“取自库存表”
    'tt = "Select * from kcJa where htBh='" & form2Htp.txtHtbh.Text & "'"
    'form2Htp.adoKu.Recordset.Close
    'form2Htp.adoKu.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'Set form2Htp.dtgKu.DataSource = form2Htp.adoKu
    
    ''打开采购表
    'ft = "Select * from CG Where htbh='" & form2Htp.txtHtbh.Text & "' and khmc<>'库存'"
    'frmAdo.adoTmp.Recordset.Close
    'frmAdo.adoTmp.Recordset.Open ft, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    'Set form2Htp.dtgCG.DataSource = frmAdo.adoTmp
    
    '打开佣金表
    tt = "Select * from Yongjin where htBh='" & form2Htp.txtHtbh.Text & "' order by yId"
    frmYj.adoYj.Recordset.Close
    frmYj.adoYj.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    Set frmYj.dtgYj.DataSource = frmYj.adoYj
    
    
    
    
    form2Htp.tabHt.TabEnabled(1) = True
    form2Htp.tabHt.TabEnabled(2) = True
    'End If
    
    
    
    
    
    
    
    form2Htp.tabHt.Tab = 0
    htBrow.MousePointer = 0
    
    
        '佣金、利润2、提成不显示
        form2Htp.txtYj1.Visible = False
        form2Htp.txtYj2.Visible = False
        form2Htp.txtLr1.Visible = False
        form2Htp.txtLr2.Visible = False
        'form2Htp.txtTc1.Visible = False
        'form2Htp.txtTc2.Visible = False
        form2Htp.lblYj.Visible = False
        form2Htp.lblLr2.Visible = False
        'form2Htp.lblTc.Visible = False
Else
        Call modHt.NewQing
        
        Call modHt.NewBound(Hid)

        frmWbNew.Visible = True

End If
End Sub

Private Sub cmdNP_Click()
Dim Pmid As Long
Dim OldPmid As Long

Dim tt As String
Dim InHtWX As Integer
Dim InHtWB As Integer
Dim InHtLP As Integer
Dim InHtCP As Integer
'Dim CHtze As Single '改单后的新金额
Dim xZ As String
On Error Resume Next
'CHtze = 0
'If mod1.PLA = False Then
'    Exit Sub
'End If


'
tt = "select pmid from pldmain where htbh='" & lblHtbh.Caption & "'"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.HTP.RecordCount > 0 Then
    MsgBox "Hello! 公司规定,一个合同只能生成一张配料单!"
    MsgBox "你别想钻空子！"
    MsgBox "怎么样，傻眼了吧 ：）"
    Exit Sub
End If
tt = "Select * from PldHt where htbh='" & lblHtbh.Caption & "'"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.HTP.RecordCount = 0 Then
    MsgBox ("信息资料有误,请与马晓聪联系!")
    Exit Sub
End If

If mod1.HTP.Fields("newF") = 1 Then

'建立新配料单
InHtWX = InStr(lblHtbh.Caption, "WX")
InHtWB = InStr(lblHtbh.Caption, "WB")
InHtLP = InStr(lblHtbh.Caption, "LP")
InHtCP = InStr(lblHtbh.Caption, "CP")

Select Case mod1.HTP.Fields("htxz").Value
Case "A. 零配件合同"
xZ = "LP"
Case "零配件"
xZ = "LP"
Case "B1.工程合同"
xZ = "GC"
Case "C. 维保合同"
xZ = "WB"
Case "维保"
xZ = "WB"
Case "D. 维修合同"
xZ = "WX"
Case "大修"
xZ = "WX"
Case "E. 产品合同"
xZ = "CP"
End Select


                 Set mod1.cmd = New ADODB.command
                 mod1.cmd.ActiveConnection = mod1.CC
                 mod1.cmd.CommandText = "PLDadd"
                 mod1.cmd.CommandType = adCmdStoredProc
                 mod1.cmd.Parameters("@htbh") = lblHtbh.Caption
                 mod1.cmd.Parameters("@xmmc") = mod1.HTP.Fields("Xmmc").Value
                 mod1.cmd.Parameters("@khdh") = mod1.HTP.Fields("Khdh").Value
                 mod1.cmd.Parameters("@htze") = mod1.HTP.Fields("htze").Value
                 mod1.cmd.Parameters("@krq") = mod1.DQda
                 mod1.cmd.Parameters("@xz") = xZ
                 mod1.cmd.Parameters("@ywy") = mod1.DName
                 mod1.cmd.Parameters("@uid") = mod1.DHid
                 mod1.cmd.Parameters("@nlb") = 64
                 mod1.cmd.Parameters("@lcou") = 6
                 mod1.cmd.Parameters("@lc") = 0
                 mod1.cmd.Parameters("@lcren") = mod1.DName
                 mod1.cmd.Parameters("@lcuid") = mod1.DHid
                 mod1.cmd.Execute
                 Pmid = mod1.cmd.Parameters("@pmid").Value
                 Set cmd = Nothing
                 

                 frmPld.Show
                 Call modPld.PLDQing
                
                 tt = "Select * from PLD where PMid=" & Pmid
                 Set mod1.HTP = New ADODB.Recordset
                 mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                 
                 '取得相应合同评审单的货品资料
                 If NewF = False Then
                    tt = "PldGxHt('" & lblHtbh.Caption & "')"
                    form2Htp.adoSale.Recordset.Close
                    form2Htp.adoSale.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
                    form2Htp.adoSale.Recordset.MoveFirst
                    Do While Not form2Htp.adoSale.Recordset.EOF
                        mod1.HTP.AddNew "htbh", form2Htp.txtHtbh.Text
                        mod1.HTP.Update "pmid", Pmid
                        mod1.HTP.Update "hpBm", form2Htp.adoSale.Recordset.Fields("hpBm").Value
                        mod1.HTP.Update "ljmc", form2Htp.adoSale.Recordset.Fields("ljmc").Value
                        mod1.HTP.Update "phBiao", form2Htp.adoSale.Recordset.Fields("phBiao").Value
                        mod1.HTP.Update "ljbh", form2Htp.adoSale.Recordset.Fields("ljbh").Value
                        mod1.HTP.Update "hplb", form2Htp.adoSale.Recordset.Fields("hplb").Value
                        mod1.HTP.Update "jldw", form2Htp.adoSale.Recordset.Fields("jldw").Value
                        mod1.HTP.Update "ljsl", form2Htp.adoSale.Recordset.Fields("ljsl").Value
                        mod1.HTP.Update "WFL", form2Htp.adoSale.Recordset.Fields("ljsl").Value
                        mod1.HTP.UpdateBatch
                        form2Htp.adoSale.Recordset.MoveNext
                    Loop
                 Else
                    tt = "PldNGxHt('" & lblHtbh.Caption & "')"
                    form2Htp.adoSale.Recordset.Close
                    form2Htp.adoSale.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
                    form2Htp.adoSale.Recordset.MoveFirst
                    Do While Not form2Htp.adoSale.Recordset.EOF
                        mod1.HTP.AddNew "htbh", lblHtbh.Caption
                        mod1.HTP.Update "pmid", Pmid
                        'mod1.HTP.Update "hpBm", form2Htp.adoSale.Recordset.Fields("hpBm").Value
                        mod1.HTP.Update "ljmc", form2Htp.adoSale.Recordset.Fields("ljmc").Value
                        If IsNull(form2Htp.adoSale.Recordset.Fields("pbcd").Value) = True Or form2Htp.adoSale.Recordset.Fields("pbcd").Value = "" Then
                            mod1.HTP.Update "phBiao", form2Htp.adoSale.Recordset.Fields("jzpb").Value
                        Else
                            mod1.HTP.Update "phBiao", form2Htp.adoSale.Recordset.Fields("pbcd").Value
                        End If
                        mod1.HTP.Update "ljbh", form2Htp.adoSale.Recordset.Fields("ljbh").Value
                        'mod1.HTP.Update "hplb", form2Htp.adoSale.Recordset.Fields("hplb").Value
                        'mod1.HTP.Update "jldw", form2Htp.adoSale.Recordset.Fields("jldw").Value
                        mod1.HTP.Update "ljsl", form2Htp.adoSale.Recordset.Fields("sl").Value
                        mod1.HTP.Update "WFL", form2Htp.adoSale.Recordset.Fields("sl").Value
                        mod1.HTP.UpdateBatch
                        form2Htp.adoSale.Recordset.MoveNext
                    Loop
                 End If
                 frmPld.lblZT.Visible = False
                 
ElseIf mod1.HTP.Fields("newF") = 2 Then
                 Set mod1.cmd = New ADODB.command
                 mod1.cmd.ActiveConnection = mod1.CC
                 mod1.cmd.CommandText = "PLDaddnew"
                 mod1.cmd.CommandType = adCmdStoredProc
                 mod1.cmd.Parameters("@htbh") = lblHtbh.Caption
                 mod1.cmd.Parameters("@hid") = mod1.HTP.Fields("hid").Value
                 mod1.cmd.Parameters("@xmmc") = mod1.HTP.Fields("Xmmc").Value
                 mod1.cmd.Parameters("@khdh") = mod1.HTP.Fields("Khdh").Value
                 mod1.cmd.Parameters("@xmadr") = mod1.HTP.Fields("xmadr").Value
                 mod1.cmd.Parameters("@htze") = mod1.HTP.Fields("htze").Value
                 mod1.cmd.Parameters("@krq") = mod1.DQda
                 mod1.cmd.Parameters("@ywy") = mod1.DName
                 mod1.cmd.Parameters("@uid") = mod1.DHid
                 mod1.cmd.Parameters("@nlb") = 64
                 mod1.cmd.Parameters("@lcou") = 6
                 mod1.cmd.Parameters("@lc") = 1
                 mod1.cmd.Parameters("@pmid") = 0
                 mod1.cmd.Parameters("@lcren") = mod1.DName
                 mod1.cmd.Parameters("@lcuid") = mod1.DHid
                 mod1.cmd.Execute
                 Pmid = mod1.cmd.Parameters("@pmid").Value
                 Set cmd = Nothing

End If
                    Call modPld.PLDQing
                    Call modPld.PLDBound(Pmid)
                    frmPld.Height = 6000
                    frmPld.cmdSave.Enabled = True

End Sub

Private Sub cmdPldOpen_Click()
Dim tt As String
Dim Pmid As Long
Dim POid As Long
On Error Resume Next
'dtgPld.Col = 2
'Pmid = dtgPld.Text
Pmid = adoPld.Fields("编号").Value
If mod1.DKZ(Pmid, 5) = True Then
        MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
        Exit Sub
End If

Call modPld.PLDQing
Call modPld.PLDBound(Pmid)

'dtgPld.Col = 4
'POid = dtgPld.Text
POid = adoPld.Fields("guid").Value
'打开旧单子
Set mod1.PldO = New ADODB.Recordset
tt = "PldOldCount(" & POid & ")"
mod1.PldO.Close
mod1.PldO.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc

If mod1.PldO.RecordCount > 0 Then
    mod1.PldO.MoveLast
    Call modPld.PldOldBound(mod1.PldO.Fields("Pmid").Value)

    frmPld.cmdRight.Enabled = False
    frmPld.cmdLeft.Enabled = True
    frmPld.Height = 9750
Else
    frmPld.Height = 5895
End If
frmPld.lblZT.Visible = True
frmPld.Visible = True
frmPld.ZOrder 0
frmHtZX.Enabled = False
End Sub


Private Sub cmdRef1_Click() '
Dim tt As String
On Error Resume Next
Select Case comXZ.Text
    Case "合同金额"
        tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where ((业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
        "')) and 合同金额=" & Val(txtYc.Text) & " and (状态='执行' or 状态='完成') order by 合同日期 desc"
    Case "项目名称"
        tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where ((业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
        "')) and 项目名称 like '%" & Trim(txtYc.Text) & "%'  and (状态='执行' or 状态='完成')  order by 合同日期 desc"
    Case "合同编号"
        tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where ((业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
        "')) and 合同编号 like '%" & Trim(txtYc.Text) & "%'  and (状态='执行' or 状态='完成')  order by 合同日期 desc"
End Select

    frmHtZX.adoBr.Close
    frmHtZX.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmHtZX.dtgBr.DataSource = frmHtZX.adoBr
    If frmHtZX.adoBr.RecordCount > 0 Then
        frmHtZX.dtgBr.FixedRows = 0
        frmHtZX.dtgBr.MergeCol(1) = True
        frmHtZX.dtgBr.MergeCol(2) = True
        frmHtZX.dtgBr.MergeCol(3) = True
        frmHtZX.dtgBr.MergeCells = 3
        frmHtZX.dtgBr.FixedRows = 1
    End If
End Sub

Private Sub Command2_Click()
Dim tt As String
On Error Resume Next
    tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where ((业务员='" & mod1.DName & "' and uid='" & mod1.DHid & "') or (xywy='" & mod1.DName & "' and Xuid='" & mod1.DHid & _
    "')) and 状态='执行' order by 合同日期 desc"
    frmHtZX.adoBr.Close
    frmHtZX.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmHtZX.dtgBr.DataSource = frmHtZX.adoBr
    If frmHtZX.adoBr.RecordCount > 0 Then
        frmHtZX.dtgBr.FixedRows = 0
        frmHtZX.dtgBr.MergeCol(1) = True
        frmHtZX.dtgBr.MergeCol(2) = True
        frmHtZX.dtgBr.MergeCol(3) = True
        frmHtZX.dtgBr.MergeCells = 3
        frmHtZX.dtgBr.FixedRows = 1
    End If
End Sub

Private Sub dtgBr_Click()
Dim tt As String
On Error Resume Next
dtgBr.Col = 7
NewF = dtgBr.Text
dtgBr.Col = 5
'If Trim(lblHtbh.Caption) <> dtgBr.Text Then
    lblHtbh.Caption = dtgBr.Text
    tt = "select * from PldView where htbh='" & lblHtbh.Caption & "' order by 编号"
    adoPld.Close
    adoPld.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set dtgPld.DataSource = adoPld
'End If
End Sub

Private Sub dtgBr_DblClick()
Static Px As Boolean

If dtgBr.Row = 1 Then
    If Px = True Then
        dtgBr.Sort = 2
        Px = False
    Else
        dtgBr.Sort = 1
        Px = True
    End If
'Else
'    MsgBox MGa.ColData(1)
End If
End Sub


Private Sub dtgBr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Zf As Boolean
If Button <> 2 Then Exit Sub
If Zf = False Then
        dtgBr.FixedRows = 0

        dtgBr.MergeCells = 0
        dtgBr.FixedRows = 1
        Zf = True
Else
        dtgBr.FixedRows = 0
        dtgBr.MergeCol(1) = True
        dtgBr.MergeCol(2) = True
        dtgBr.MergeCol(3) = True
        dtgBr.MergeCells = 3
        dtgBr.FixedRows = 1
        Zf = False
End If
End Sub

Private Sub dtgBr_RowColChange()
Dim tt As String
On Error Resume Next
dtgBr.Col = 7
NewF = dtgBr.Text
dtgBr.Col = 5
If Trim(lblHtbh.Caption) <> dtgBr.Text Then
    lblHtbh.Caption = dtgBr.Text
    tt = "select * from PldView where htbh='" & lblHtbh.Caption & "' order by 编号"
    adoPld.Close
    adoPld.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set dtgPld.DataSource = adoPld
End If
End Sub


Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
Set adoBr = New ADODB.Recordset
Set adoPld = New ADODB.Recordset
dtgBr.ColWidth(0) = 300
frmHtZX.dtgBr.ColWidth(1) = 3000
frmHtZX.dtgBr.ColWidth(3) = 1300
frmHtZX.dtgBr.ColWidth(5) = 1800
dtgBr.ColWidth(6) = 0
dtgBr.ColWidth(7) = 0

'dtgPld.ColWidth(0) = 300
'dtgPld.ColWidth(3) = 3200
'dtgPld.ColWidth(4) = 0
'dtgPld.ColWidth(5) = 0
'dtgPld.ColWidth(7) = 0
'dtgPld.ColWidth(8) = 0
'dtgPld.ColWidth(9) = 0
End Sub
