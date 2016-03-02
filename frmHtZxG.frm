VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmHtZxG 
   Caption         =   "合同执行列表"
   ClientHeight    =   9135
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15210
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdFw 
      Caption         =   "查询范围"
      Height          =   315
      Left            =   9120
      TabIndex        =   16
      Top             =   8640
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "导航"
      Height          =   585
      Left            =   14520
      Picture         =   "frmHtZxG.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8580
      Width           =   675
   End
   Begin VB.Frame Frame2 
      Caption         =   "条件查询"
      Height          =   615
      Left            =   30
      TabIndex        =   8
      Top             =   8520
      Width           =   5745
      Begin VB.TextBox txtYc 
         Height          =   285
         Left            =   2820
         TabIndex        =   11
         Top             =   240
         Width           =   1635
      End
      Begin VB.ComboBox comXZ 
         Height          =   300
         ItemData        =   "frmHtZxG.frx":0102
         Left            =   810
         List            =   "frmHtZxG.frx":010F
         TabIndex        =   10
         Text            =   "合同金额"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdRef1 
         Caption         =   "查  询"
         Height          =   285
         Left            =   4590
         TabIndex        =   9
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "值"
         Height          =   255
         Left            =   2610
         TabIndex        =   13
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "条件"
         Height          =   255
         Left            =   300
         TabIndex        =   12
         Top             =   300
         Width           =   795
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "所有执行合同"
      Height          =   375
      Left            =   5820
      TabIndex        =   7
      Top             =   8730
      Width           =   1515
   End
   Begin VB.CommandButton cmdHtOpen 
      Caption         =   "打开"
      Height          =   345
      Left            =   7440
      TabIndex        =   6
      Top             =   8730
      Width           =   1365
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
      Height          =   2895
      Left            =   9060
      TabIndex        =   1
      Top             =   30
      Width           =   6135
      Begin VB.CommandButton cmdAll 
         Caption         =   "所  有"
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   2280
         Width           =   2985
      End
      Begin VB.CommandButton cmdNP 
         Caption         =   "新建配料单"
         Height          =   315
         Left            =   3120
         TabIndex        =   3
         Top             =   2550
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton cmdPldOpen 
         Caption         =   "打  开"
         Height          =   315
         Left            =   4620
         TabIndex        =   2
         Top             =   2550
         Width           =   1485
      End
      Begin MSDataGridLib.DataGrid dtgPld 
         Height          =   1995
         Left            =   0
         TabIndex        =   19
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3519
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
      Begin VB.Label Label3 
         Caption         =   "合同编号"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   2370
         Width           =   735
      End
      Begin VB.Label lblHtbh 
         Height          =   225
         Left            =   930
         TabIndex        =   4
         Top             =   2370
         Width           =   2265
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "财务评定"
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
      Height          =   5505
      Left            =   9030
      TabIndex        =   0
      Top             =   2940
      Width           =   6165
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8445
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   14896
      _Version        =   393216
      BackColor       =   -2147483634
      BackColorBkg    =   -2147483636
      FillStyle       =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblFw 
      Height          =   285
      Left            =   10290
      TabIndex        =   17
      Top             =   8670
      Width           =   2475
   End
End
Attribute VB_Name = "frmHtZxG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoBr As ADODB.Recordset
Public adoPld As ADODB.Recordset
Dim NewF As Boolean

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
End Sub

Private Sub cmdFw_Click()
Set Ren.XForm = New frmHtZxG
Call mod1.RenXz("frmHtZxG", Me, 0)
End Sub

Private Sub cmdHtOpen_Click()

Dim tt As String
Dim xZ As String

Dim Hid As Long
'Dim Lid As String
On Error Resume Next

If mod1.BM = "杰升" And mod1.DName <> "罗黎" Then
    MsgBox ("不许偷看！")
    Exit Sub
End If
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
If mod1.PLA = False Then
    Exit Sub
End If

'DD = InputBox("请键入合同编号,以关联正确的合同评审单")
'If DD = "" Then
'    Exit Sub
'End If
'
tt = "Select * from PldHt where htbh='" & lblHtbh.Caption & "'"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'CHtze = mod1.HTP.Fields("htze").Value
'
'If mod1.HTP.RecordCount = 1 Then
'    If mod1.HTP.Fields("htF") = 2 Then
'        MsgBox ("此合同已经完成,不能再配料!")
'        Exit Sub
'    ElseIf mod1.HTP.Fields("htF").Value <> 1 Then
'        MsgBox ("此合同未被执行,不能配料!")
'        Exit Sub
'    End If

'DD = UCase(DD)
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
                    Call modPld.PLDQing
                    Call modPld.PLDBound(Pmid)
                    frmPld.Height = 6000
'        Else ' 有重复单子
''                If (InHtWX > 3) Or InHtWB > 0 Then '如果是维保维修单子,则不能新建第二张配料单
''                    MsgBox "有相同配料单,所以不能新建!"
''                    Exit Sub
''                End If
'
'                If mod1.DKZ(mod1.PldV.Fields(mod1.HTT.Fields("pmid").Value).Value, 5) = True Then
'                        MsgBox "有相同配料单,由于这份表单正由" & mod1.DKRen & "打开,所以无法改单,请稍候再试,或与马晓聪联系."
'                        Exit Sub
'                End If
'                Set mod1.CMD = New ADODB.command
'                mod1.CMD.ActiveConnection = mod1.CC
'                mod1.CMD.CommandText = "PLDgd"
'                mod1.CMD.CommandType = adCmdStoredProc
'                mod1.CMD.Parameters("@htbh") = DD
'                mod1.CMD.Parameters("@xmmc") = mod1.HTT.Fields("xmmc").Value
'                mod1.CMD.Parameters("@htze") = CHtze
'                mod1.CMD.Parameters("@krq") = mod1.DQda
'                mod1.CMD.Parameters("@xmADR") = mod1.HTT.Fields("xmAdr").Value
'                mod1.CMD.Parameters("@Pmid") = mod1.HTT.Fields("pmid").Value
'                mod1.CMD.Parameters("@Guid") = mod1.HTT.Fields("guid").Value
'                mod1.CMD.Execute
'                Set CMD = Nothing
'
'                    tt = "select pmid from maxPld"
'                    Set mod1.HTP = New ADODB.Recordset
'                    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'                    Pmid = mod1.HTP.Fields("pmid").Value
'
'                 '取得相应合同评审单的货品资料
'                 tt = "PldGxHt('" & DD & "')"
'                 form2Htp.adoSale.Recordset.Close
'                 form2Htp.adoSale.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
'                 tt = " PLDBoundB('" & Pmid & "')"
'                 frmPld.adoHp.Recordset.Close
'                 frmPld.adoHp.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
'                 form2Htp.adoSale.Recordset.MoveFirst
'                 Do While Not form2Htp.adoSale.Recordset.EOF
'                     frmPld.adoHp.Recordset.AddNew "htbh", form2Htp.txtHtbh.Text
'                     frmPld.adoHp.Recordset.Update "pmid", Pmid
'                     frmPld.adoHp.Recordset.Update "hpBm", form2Htp.adoSale.Recordset.Fields("hpBm").Value
'                     frmPld.adoHp.Recordset.Update "ljmc", form2Htp.adoSale.Recordset.Fields("ljmc").Value
'                     frmPld.adoHp.Recordset.Update "phBiao", form2Htp.adoSale.Recordset.Fields("phBiao").Value
'                     frmPld.adoHp.Recordset.Update "ljbh", form2Htp.adoSale.Recordset.Fields("ljbh").Value
'                     frmPld.adoHp.Recordset.Update "hplb", form2Htp.adoSale.Recordset.Fields("hplb").Value
'                     frmPld.adoHp.Recordset.Update "jldw", form2Htp.adoSale.Recordset.Fields("jldw").Value
'                     frmPld.adoHp.Recordset.Update "ljsl", form2Htp.adoSale.Recordset.Fields("ljsl").Value
'                     frmPld.adoHp.Recordset.Update "WFL", form2Htp.adoSale.Recordset.Fields("ljsl").Value
'                     frmPld.adoHp.Recordset.UpdateBatch
'                     form2Htp.adoSale.Recordset.MoveNext
'                 Loop
'                 frmPld.lblZT.Visible = False
'
'                    '更新当前配料单
'                    OldPmid = mod1.HTT.Fields("pmid").Value
'
'
'                    Call modPld.PLDQing
'                    Call modPld.PLDBound(Pmid)
'                    Call modPld.PldOldBound(OldPmid)
'
'
'
'                    '刷新旧单列表
'                    tt = "PldOldCount(" & frmPld.lblGuid.Caption & ")"
'                    mod1.PldO.Close
'                    mod1.PldO.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
'                    mod1.PldO.MoveLast
'                    frmPld.cmdRight.Enabled = False
'
'                    frmPld.dtgSale.Columns("产品名称").Locked = False
'                    frmPld.dtgSale.Columns("牌号商标").Locked = False
'                    frmPld.dtgSale.Columns("规格型号").Locked = False
'                    frmPld.dtgSale.Columns("单位").Locked = False
'                    frmPld.dtgSale.Columns("数量").Locked = False
'                    frmPld.cmdAD.Visible = True
'                    frmPld.cmdDE.Visible = True
'                    frmPld.cmdSave.Enabled = True
'                    frmPld.Height = 10305
'                    frmPld.cmdRight.Enabled = True
'
'        End If
        

'    End If
'Else
'    MsgBox ("您输入的合同编号有误,请仔细核对!")
'End If
End Sub

Private Sub cmdPldOpen_Click()
Dim tt As String
Dim Pmid As Long
Dim POid As Long
On Error Resume Next
'dtgPld.Col = 2
Pmid = adoPld.Fields("编号").Value
Pmid = dtgPld.Text
If mod1.DKZ(Pmid, 5) = True Then
        MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
        Exit Sub
End If

Call modPld.PLDQing
Call modPld.PLDBound(Pmid)

dtgPld.Col = 4
POid = dtgPld.Text
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
frmHtZxG.Enabled = False
End Sub


Private Sub cmdRef1_Click()
Dim tt As String
On Error Resume Next
Select Case comXZ.Text
    Case "合同金额"
        If mod1.KhK = 1 And mod1.BmJl = True Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where 部门='" & mod1.BM & "' and 合同金额=" & Val(txtYc.Text) & " and 状态='执行' order by 合同日期 desc"
        ElseIf mod1.KhK = 2 Or mod1.BM = "财务部" Then
            If mod1.Qy = "上海" Then
                tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where comid=0 and 合同金额=" & Val(txtYc.Text) & _
                " and 状态='执行' and not(部门='维销部3' or 部门='产品部1' or 部门='产品部2')  order by 合同日期 desc"
            ElseIf mod1.Qy = "广州" Then
                tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where comid=1 and 合同金额=" & Val(txtYc.Text) & " and 状态='执行' order by 合同日期 desc"
            End If
        ElseIf (mod1.VLP = 2 Or mod1.VLP = 3 Or mod1.DName = "罗黎" Or mod1.DName = "钱之文") And mod1.Bq2 = False Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where 合同金额=" & Val(txtYc.Text) & " and 状态='执行' order by 合同日期 desc"
        ElseIf mod1.KhK = 3 Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where 合同金额=" & Val(txtYc.Text) & " and 状态='执行' order by 合同日期 desc"
        ElseIf mod1.Bq2 = True And mod1.Qy <> "上海" Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where 区域='" & mod1.Qy & "' and 合同金额=" & Val(txtYc.Text) & " and 状态='执行' order by 合同日期 desc"
        End If
        
    Case "项目名称"
        If mod1.KhK = 1 And mod1.BmJl = True Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where  部门='" & mod1.BM & "'  and 项目名称 like '%" & Trim(txtYc.Text) & "%'  and 状态='执行'  order by 合同日期 desc"
        ElseIf mod1.KhK = 2 Or mod1.BM = "财务部" Then
            If mod1.Qy = "上海" Then
                tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where  comid=0  and 项目名称 like '%" & Trim(txtYc.Text) & _
                "%'  and 状态='执行'  and not(部门='维销部3' or 部门='产品部1' or 部门='产品部2')   order by 合同日期 desc"
            ElseIf mod1.Qy = "广州" Then
                tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where  comid=1  and 项目名称 like '%" & Trim(txtYc.Text) & "%'  and 状态='执行'  order by 合同日期 desc"
            End If
        ElseIf (mod1.VLP = 2 Or mod1.VLP = 3 Or mod1.DName = "罗黎" Or mod1.DName = "钱之文") And mod1.Bq2 = False Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where 项目名称 like '%" & Trim(txtYc.Text) & "%'  and 状态='执行'  order by 合同日期 desc"
        ElseIf mod1.KhK = 3 Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where 项目名称 like '%" & Trim(txtYc.Text) & "%'  and 状态='执行'  order by 合同日期 desc"
        ElseIf mod1.Bq2 = True And mod1.Qy <> "上海" Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where  区域='" & mod1.Qy & "'  and 项目名称 like '%" & Trim(txtYc.Text) & "%'  and 状态='执行'  order by 合同日期 desc"
        End If
    Case "合同编号"
        If mod1.KhK = 1 And mod1.BmJl = True Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where  部门='" & mod1.BM & "' and 合同编号 like '%" & Trim(txtYc.Text) & "%'  and 状态='执行'  order by 合同日期 desc"
        ElseIf mod1.KhK = 2 Or mod1.BM = "财务部" Then
            If mod1.Qy = "上海" Then
                tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where comid=0 and 合同编号 like '%" & Trim(txtYc.Text) & _
                "%'  and 状态='执行' and not(部门='维销部3' or 部门='产品部1' or 部门='产品部2')   order by 合同日期 desc"
            ElseIf mod1.Qy = "广州" Then
                tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where comid=1 and 合同编号 like '%" & Trim(txtYc.Text) & "%'  and 状态='执行'  order by 合同日期 desc"
            End If
        ElseIf (mod1.VLP = 2 Or mod1.VLP = 3 Or mod1.DName = "罗黎" Or mod1.DName = "钱之文") And mod1.Bq2 = False Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where 合同编号 like '%" & Trim(txtYc.Text) & "%'  and 状态='执行'  order by 合同日期 desc"
        ElseIf mod1.KhK = 3 Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where 合同编号 like '%" & Trim(txtYc.Text) & "%'  and 状态='执行'  order by 合同日期 desc"
        ElseIf mod1.Bq2 = True And mod1.Qy <> "上海" Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where  区域='" & mod1.Qy & "' and 合同编号 like '%" & Trim(txtYc.Text) & "%'  and 状态='执行'  order by 合同日期 desc"
        End If
End Select

    frmHtZxG.adoBr.Close
    frmHtZxG.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmHtZxG.dtgBr.DataSource = frmHtZxG.adoBr
    If frmHtZxG.adoBr.RecordCount > 0 Then
        frmHtZxG.dtgBr.FixedRows = 0
        frmHtZxG.dtgBr.MergeCol(1) = True
        frmHtZxG.dtgBr.MergeCol(2) = True
        frmHtZxG.dtgBr.MergeCol(3) = True
        frmHtZxG.dtgBr.MergeCol(8) = True
        frmHtZxG.dtgBr.MergeCells = 3
        frmHtZxG.dtgBr.FixedRows = 1
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
Dim tt As String
On Error Resume Next
        If mod1.KhK = 1 And mod1.BmJl = True Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where 部门='" & mod1.BM & "'  and 状态='执行' order by 合同日期 desc"
        ElseIf mod1.KhK = 2 Then
            If mod1.Qy = "上海" Then
                tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where comid=0  and 状态='执行'  and not(部门='维销部3' or 部门='产品部1' or 部门='产品部2')  order by 合同日期 desc"
            ElseIf mod1.Qy = "广州" Then
                tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where comid=1  and 状态='执行' order by 合同日期 desc"
            End If
        ElseIf (mod1.VLP = 2 Or mod1.VLP = 3 Or mod1.DName = "冯建川") And mod1.KhK <> 3 Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where  状态='执行' order by 合同日期 desc"
        ElseIf mod1.KhK = 3 Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where  状态='执行' and comid=" & mod1.comId & " order by 合同日期 desc"
        ElseIf mod1.Bq2 = True And mod1.Qy <> "上海" Then
            tt = "Select 项目名称,合同日期,合同性质,合同金额,合同编号,Hid,newF from htView where 区域='" & mod1.Qy & "'  and 状态='执行' order by 合同日期 desc"
        End If
    frmHtZxG.adoBr.Close
    frmHtZxG.adoBr.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmHtZxG.dtgBr.DataSource = frmHtZxG.adoBr
    If frmHtZxG.adoBr.RecordCount > 0 Then
        frmHtZxG.dtgBr.FixedRows = 0
        frmHtZxG.dtgBr.MergeCol(1) = True
        frmHtZxG.dtgBr.MergeCol(2) = True
        frmHtZxG.dtgBr.MergeCol(3) = True
        frmHtZxG.dtgBr.MergeCells = 3
        frmHtZxG.dtgBr.FixedRows = 1
    End If
End Sub

Private Sub dtgBr_Click()
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
frmHtZxG.dtgBr.ColWidth(1) = 3000
frmHtZxG.dtgBr.ColWidth(3) = 1300
frmHtZxG.dtgBr.ColWidth(5) = 1800
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


