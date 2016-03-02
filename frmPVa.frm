VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPVa 
   Caption         =   "配料单查询"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   7455
   Begin VB.Frame frmF 
      Height          =   3105
      Left            =   2100
      TabIndex        =   33
      Top             =   390
      Visible         =   0   'False
      Width           =   1275
      Begin VB.OptionButton optP 
         Caption         =   "新单"
         Height          =   255
         Left            =   60
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optQ 
         Caption         =   "执行"
         Height          =   315
         Left            =   60
         TabIndex        =   36
         Top             =   780
         Width           =   945
      End
      Begin VB.OptionButton optR 
         Caption         =   "完成"
         Height          =   315
         Left            =   60
         TabIndex        =   35
         Top             =   1290
         Width           =   1005
      End
      Begin VB.OptionButton optZf 
         Caption         =   "作废"
         Height          =   225
         Left            =   60
         TabIndex        =   34
         Top             =   1830
         Width           =   885
      End
   End
   Begin VB.TextBox txtBr 
      Height          =   285
      Left            =   2820
      TabIndex        =   26
      Top             =   4020
      Width           =   2805
   End
   Begin VB.ComboBox comBR 
      Height          =   300
      ItemData        =   "frmPVa.frx":0000
      Left            =   1350
      List            =   "frmPVa.frx":000D
      TabIndex        =   25
      Text            =   "合同金额"
      Top             =   4020
      Width           =   1485
   End
   Begin VB.Frame frmE 
      Height          =   3105
      Left            =   150
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   1275
      Begin VB.OptionButton optZe 
         Caption         =   "作废"
         Height          =   225
         Left            =   60
         TabIndex        =   32
         Top             =   1830
         Width           =   885
      End
      Begin VB.OptionButton optJ 
         Caption         =   "未结帐"
         Height          =   255
         Left            =   60
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optK 
         Caption         =   "执行"
         Height          =   315
         Left            =   60
         TabIndex        =   22
         Top             =   780
         Width           =   945
      End
      Begin VB.OptionButton optL 
         Caption         =   "完成"
         Height          =   315
         Left            =   60
         TabIndex        =   21
         Top             =   1290
         Width           =   1005
      End
   End
   Begin VB.Frame frmD 
      Height          =   3105
      Left            =   1500
      TabIndex        =   16
      Top             =   660
      Visible         =   0   'False
      Width           =   1275
      Begin VB.OptionButton optZd 
         Caption         =   "作废"
         Height          =   225
         Left            =   60
         TabIndex        =   31
         Top             =   1860
         Width           =   885
      End
      Begin VB.OptionButton optG 
         Caption         =   "未发货"
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optH 
         Caption         =   "执行"
         Height          =   315
         Left            =   60
         TabIndex        =   18
         Top             =   780
         Width           =   945
      End
      Begin VB.OptionButton optI 
         Caption         =   "完成"
         Height          =   315
         Left            =   60
         TabIndex        =   17
         Top             =   1290
         Width           =   1005
      End
   End
   Begin VB.Frame frmC 
      Height          =   3105
      Left            =   2820
      TabIndex        =   12
      Top             =   570
      Visible         =   0   'False
      Width           =   1275
      Begin VB.OptionButton optZc 
         Caption         =   "作废"
         Height          =   225
         Left            =   60
         TabIndex        =   30
         Top             =   1860
         Width           =   885
      End
      Begin VB.OptionButton optD 
         Caption         =   "未确认"
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optE 
         Caption         =   "执行"
         Height          =   315
         Left            =   60
         TabIndex        =   14
         Top             =   780
         Width           =   945
      End
      Begin VB.OptionButton optF 
         Caption         =   "完成"
         Height          =   315
         Left            =   60
         TabIndex        =   13
         Top             =   1290
         Width           =   1005
      End
   End
   Begin VB.Frame frmB 
      Height          =   3105
      Left            =   4260
      TabIndex        =   8
      Top             =   450
      Visible         =   0   'False
      Width           =   1275
      Begin VB.OptionButton optZb 
         Caption         =   "作废"
         Height          =   225
         Left            =   60
         TabIndex        =   29
         Top             =   1830
         Width           =   885
      End
      Begin VB.CommandButton cmdAddKc 
         Caption         =   "新建库存单"
         Height          =   315
         Left            =   30
         TabIndex        =   27
         Top             =   2760
         Width           =   1185
      End
      Begin VB.OptionButton optM 
         Caption         =   "未处理"
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optN 
         Caption         =   "执行"
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   780
         Width           =   945
      End
      Begin VB.OptionButton optO 
         Caption         =   "完成"
         Height          =   315
         Left            =   60
         TabIndex        =   9
         Top             =   1290
         Width           =   1005
      End
   End
   Begin VB.Frame frmA 
      Height          =   3105
      Left            =   6120
      TabIndex        =   3
      Top             =   1260
      Visible         =   0   'False
      Width           =   1275
      Begin VB.OptionButton optZa 
         Caption         =   "作废"
         Height          =   225
         Left            =   60
         TabIndex        =   28
         Top             =   1830
         Width           =   885
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "新建配料单"
         Height          =   345
         Left            =   30
         TabIndex        =   7
         Top             =   2730
         Width           =   1215
      End
      Begin VB.OptionButton optC 
         Caption         =   "完成"
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   1290
         Width           =   1005
      End
      Begin VB.OptionButton optB 
         Caption         =   "执行"
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   780
         Width           =   945
      End
      Begin VB.OptionButton optA 
         Caption         =   "新单"
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdRef 
      Caption         =   "刷新"
      Height          =   285
      Left            =   6210
      TabIndex        =   2
      Top             =   570
      Width           =   1005
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "打开"
      Height          =   285
      Left            =   6210
      TabIndex        =   1
      Top             =   210
      Width           =   1005
   End
   Begin MSDataGridLib.DataGrid dtgPld 
      Height          =   3945
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   6959
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "XMMC"
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
         DataField       =   "htZe"
         Caption         =   "合同金额"
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
         DataField       =   "KRQ"
         Caption         =   "开单日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "yyyy-M-d"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   3
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2294.929
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "查询方式:"
      Height          =   225
      Left            =   330
      TabIndex        =   24
      Top             =   4080
      Width           =   1065
   End
End
Attribute VB_Name = "frmPVa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAdd_Click()
Dim Pmid As Long
Dim OldPmid As Long
Dim DD As String
Dim tt As String
Dim InHtWX As Integer
Dim InHtWB As Integer
Dim InHtLP As Integer
Dim InHtCP As Integer
Dim CHtze As Single '改单后的新金额
Dim xZ As String
On Error Resume Next
CHtze = 0
If mod1.PLA = False Then
    Exit Sub
End If

DD = InputBox("请键入合同编号,以关联正确的合同评审单")
If DD = "" Then
    Exit Sub
End If

tt = "Select * from PldHt where htbh='" & DD & "'"
mod1.HtP.Close
mod1.HtP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
CHtze = mod1.HtP.Fields("htze").Value

If mod1.HtP.RecordCount = 1 Then
    If mod1.HtP.Fields("htF") = 2 Then
        MsgBox ("此合同已经完成,不能再配料!")
        Exit Sub
    ElseIf mod1.HtP.Fields("htF").Value <> 1 Then
        MsgBox ("此合同未被执行,不能配料!")
        Exit Sub
    End If

DD = UCase(DD)
'建立新配料单
InHtWX = InStr(DD, "WX")
InHtWB = InStr(DD, "WB")
InHtLP = InStr(DD, "LP")
InHtCP = InStr(DD, "CP")

Select Case mod1.HtP.Fields("htxz").Value
Case "A. 零配件合同"
xZ = "LP"
Case "B1.工程合同"
xZ = "GC"
Case "C. 维保合同"
xZ = "WB"
Case "D. 维修合同"
xZ = "WX"
Case "E. 产品合同"
xZ = "CP"
End Select

''维保维修合同
'    If (InHtWX > 3) Or InHtWB > 0 Then
'        Set mod1.CMD = New ADODB.Command
'        mod1.CMD.ActiveConnection = mod1.CC
'        mod1.CMD.CommandText = "PLDadd"
'        mod1.CMD.CommandType = adCmdStoredProc
'        mod1.CMD.Parameters("@htbh") = DD
'        mod1.CMD.Parameters("@xmmc") = mod1.HtP.Fields("Xmmc").Value
'        mod1.CMD.Parameters("@khdh") = mod1.HtP.Fields("Khdh").Value
'        mod1.CMD.Parameters("@htze") = mod1.HtP.Fields("htze").Value
'        mod1.CMD.Parameters("@krq") = mod1.DQda
'        mod1.CMD.Execute
'        Set CMD = Nothing
'
'        tt = "select pmid from maxPld"
'        mod1.HtP.Close
'        mod1.HtP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
'        Pmid = mod1.HtP.Fields("pmid").Value
'        frmPld.Show
'        Call modPld.PLDQing
'        Call modPld.PLDBound(Pmid)
'        frmPld.lblZT.Visible = False
'        frmPld.Height = 6000
'    ElseIf InHtLP > 0 Or InHtCP > 0 Then '零配件合同
        ' 检查有无重复单子
        tt = "pldjc('" & DD & "')"
        mod1.HtT.Close
        mod1.HtT.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        If mod1.HtT.RecordCount = 0 Then
                 Set mod1.CMD = New ADODB.Command
                 mod1.CMD.ActiveConnection = mod1.CC
                 mod1.CMD.CommandText = "PLDadd"
                 mod1.CMD.CommandType = adCmdStoredProc
                 mod1.CMD.Parameters("@htbh") = DD
                 mod1.CMD.Parameters("@xmmc") = mod1.HtP.Fields("Xmmc").Value
                 mod1.CMD.Parameters("@khdh") = mod1.HtP.Fields("Khdh").Value
                 mod1.CMD.Parameters("@htze") = mod1.HtP.Fields("htze").Value
                 mod1.CMD.Parameters("@krq") = mod1.DQda
                 mod1.CMD.Parameters("@xz") = xZ
                 mod1.CMD.Execute
                 Set CMD = Nothing
                 
                 tt = "select pmid from maxPld"
                 mod1.HtP.Close
                 mod1.HtP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
                 Pmid = mod1.HtP.Fields("pmid").Value
                 frmPld.Show
                 Call modPld.PLDQing
                
                 tt = "Select * from PLD where PMid=" & Pmid
                 mod1.HtP.Close
                 mod1.HtP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
                 
                 '取得相应合同评审单的货品资料
                 tt = "PldGxHt('" & DD & "')"
                 form2Htp.adoSale.Recordset.Close
                 form2Htp.adoSale.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
                 form2Htp.adoSale.Recordset.MoveFirst
                 Do While Not form2Htp.adoSale.Recordset.EOF
                     mod1.HtP.AddNew "htbh", form2Htp.txtHtbh.Text
                     mod1.HtP.Update "pmid", Pmid
                     mod1.HtP.Update "hpBm", form2Htp.adoSale.Recordset.Fields("hpBm").Value
                     mod1.HtP.Update "ljmc", form2Htp.adoSale.Recordset.Fields("ljmc").Value
                     mod1.HtP.Update "phBiao", form2Htp.adoSale.Recordset.Fields("phBiao").Value
                     mod1.HtP.Update "ljbh", form2Htp.adoSale.Recordset.Fields("ljbh").Value
                     mod1.HtP.Update "hplb", form2Htp.adoSale.Recordset.Fields("hplb").Value
                     mod1.HtP.Update "jldw", form2Htp.adoSale.Recordset.Fields("jldw").Value
                     mod1.HtP.Update "ljsl", form2Htp.adoSale.Recordset.Fields("ljsl").Value
                     mod1.HtP.Update "WFL", form2Htp.adoSale.Recordset.Fields("ljsl").Value
                     mod1.HtP.UpdateBatch
                     form2Htp.adoSale.Recordset.MoveNext
                 Loop
                 frmPld.lblZT.Visible = False
                    Call modPld.PLDQing
                    Call modPld.PLDBound(Pmid)
                    frmPld.Height = 6000
        Else ' 有重复单子
                If (InHtWX > 3) Or InHtWB > 0 Then '如果是维保维修单子,则不能新建第二张配料单
                    MsgBox "有相同配料单,所以不能新建!"
                    Exit Sub
                End If
        
                If mod1.DKZ(mod1.PldV.Fields(mod1.HtT.Fields("pmid").Value).Value, 5) = True Then
                        MsgBox "有相同配料单,由于这份表单正由" & mod1.DKRen & "打开,所以无法改单,请稍候再试,或与马晓聪联系."
                        Exit Sub
                End If
                Set mod1.CMD = New ADODB.Command
                mod1.CMD.ActiveConnection = mod1.CC
                mod1.CMD.CommandText = "PLDgd"
                mod1.CMD.CommandType = adCmdStoredProc
                mod1.CMD.Parameters("@htbh") = DD
                mod1.CMD.Parameters("@xmmc") = mod1.HtT.Fields("xmmc").Value
                mod1.CMD.Parameters("@htze") = CHtze
                mod1.CMD.Parameters("@krq") = mod1.DQda
                mod1.CMD.Parameters("@xmADR") = mod1.HtT.Fields("xmAdr").Value
                mod1.CMD.Parameters("@Pmid") = mod1.HtT.Fields("pmid").Value
                mod1.CMD.Parameters("@Guid") = mod1.HtT.Fields("guid").Value
                mod1.CMD.Execute
                Set CMD = Nothing
                
                    tt = "select pmid from maxPld"
                    mod1.HtP.Close
                    mod1.HtP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
                    Pmid = mod1.HtP.Fields("pmid").Value
                
                 '取得相应合同评审单的货品资料
                 tt = "PldGxHt('" & DD & "')"
                 form2Htp.adoSale.Recordset.Close
                 form2Htp.adoSale.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
                 tt = " PLDBoundB('" & Pmid & "')"
                 frmPld.adoHp.Recordset.Close
                 frmPld.adoHp.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
                 form2Htp.adoSale.Recordset.MoveFirst
                 Do While Not form2Htp.adoSale.Recordset.EOF
                     frmPld.adoHp.Recordset.AddNew "htbh", form2Htp.txtHtbh.Text
                     frmPld.adoHp.Recordset.Update "pmid", Pmid
                     frmPld.adoHp.Recordset.Update "hpBm", form2Htp.adoSale.Recordset.Fields("hpBm").Value
                     frmPld.adoHp.Recordset.Update "ljmc", form2Htp.adoSale.Recordset.Fields("ljmc").Value
                     frmPld.adoHp.Recordset.Update "phBiao", form2Htp.adoSale.Recordset.Fields("phBiao").Value
                     frmPld.adoHp.Recordset.Update "ljbh", form2Htp.adoSale.Recordset.Fields("ljbh").Value
                     frmPld.adoHp.Recordset.Update "hplb", form2Htp.adoSale.Recordset.Fields("hplb").Value
                     frmPld.adoHp.Recordset.Update "jldw", form2Htp.adoSale.Recordset.Fields("jldw").Value
                     frmPld.adoHp.Recordset.Update "ljsl", form2Htp.adoSale.Recordset.Fields("ljsl").Value
                     frmPld.adoHp.Recordset.Update "WFL", form2Htp.adoSale.Recordset.Fields("ljsl").Value
                     frmPld.adoHp.Recordset.UpdateBatch
                     form2Htp.adoSale.Recordset.MoveNext
                 Loop
                 frmPld.lblZT.Visible = False
                
                    '更新当前配料单
                    OldPmid = mod1.HtT.Fields("pmid").Value
                    

                    Call modPld.PLDQing
                    Call modPld.PLDBound(Pmid)
                    Call modPld.PldOldBound(OldPmid)
                    
                    
                    
                    '刷新旧单列表
                    tt = "PldOldCount(" & frmPld.lblGuid.Caption & ")"
                    mod1.PldO.Close
                    mod1.PldO.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
                    mod1.PldO.MoveLast
                    frmPld.cmdRight.Enabled = False
                
                    frmPld.dtgSale.Columns("产品名称").Locked = False
                    frmPld.dtgSale.Columns("牌号商标").Locked = False
                    frmPld.dtgSale.Columns("规格型号").Locked = False
                    frmPld.dtgSale.Columns("单位").Locked = False
                    frmPld.dtgSale.Columns("数量").Locked = False
                    frmPld.cmdAD.Visible = True
                    frmPld.cmdDE.Visible = True
                    frmPld.cmdSave.Enabled = True
                    frmPld.Height = 10305
                    frmPld.cmdRight.Enabled = True
        
        End If
        

'    End If
Else
    MsgBox ("您输入的合同编号有误,请仔细核对!")
End If
    
End Sub



















Private Sub cmdAddKc_Click()
Dim Pmid As Long
Dim DD As String
Dim tt As String
On Error Resume Next
If mod1.PLB = False Then
    Exit Sub
End If
DD = "上海库存"


'建立新配料单
    Set mod1.CMD = New ADODB.Command
    mod1.CMD.ActiveConnection = mod1.CC
    mod1.CMD.CommandText = "PLDadd"
    mod1.CMD.CommandType = adCmdStoredProc
    mod1.CMD.Parameters("@htbh") = DD
    mod1.CMD.Parameters("@xmmc") = DD
    mod1.CMD.Parameters("@khdh") = "8888888"
    mod1.CMD.Parameters("@htze") = 0
    mod1.CMD.Parameters("@krq") = mod1.DQda
    mod1.CMD.Execute
    Set CMD = Nothing
    
    tt = "select pmid from maxPld"
    mod1.HtP.Close
    mod1.HtP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Pmid = mod1.HtP.Fields("pmid").Value
    frmPld.Show
    
    Call modPld.PLDQing
    Call modPld.PLDBound(Pmid)
    frmPld.lblZT.Visible = False
    frmPld.txtKhAdr.Text = "东方路3601号B座1层"
    frmPld.dtgSale.Columns("产品名称").Locked = False
    frmPld.dtgSale.Columns("牌号商标").Locked = False
    frmPld.dtgSale.Columns("规格型号").Locked = False
    frmPld.dtgSale.Columns("单位").Locked = False
    frmPld.dtgSale.Columns("数量").Locked = False
    frmPld.cmdAD.Visible = True
    frmPld.cmdDE.Visible = True
    
End Sub


Private Sub cmdOpen_Click()
Dim tt As String
On Error Resume Next
If mod1.DKZ(mod1.PldV.Fields("PMid").Value, 5) = True Then
        MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
        Exit Sub
End If

Call modPld.PLDQing
Call modPld.PLDBound(mod1.PldV.Fields("Pmid").Value)

'打开旧单子
Set mod1.PldO = New ADODB.Recordset
tt = "PldOldCount(" & mod1.PldV.Fields("Guid").Value & ")"
mod1.PldO.Close
mod1.PldO.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc

If mod1.PldO.RecordCount > 0 Then
    mod1.PldO.MoveLast
    Call modPld.PldOldBound(mod1.PldO.Fields("Pmid").Value)

    frmPld.cmdRight.Enabled = False
    frmPld.cmdLeft.Enabled = True
    frmPld.Height = 10305
Else
    frmPld.Height = 6000
End If
frmPld.lblZT.Visible = True
End Sub

Private Sub cmdReF_Click()
Dim tt As String
On Error Resume Next
If opTa.Value = True Then
    tt = "Select * from PldnewVa"
ElseIf opTB.Value = True Then
    tt = "Select * from PldnewVb"
ElseIf opTC.Value = True Then
    tt = "Select * from PldnewVc"
ElseIf optD.Value = True Then
    tt = "Select * from PldnewVac"
ElseIf optE.Value = True Then
    tt = "Select * from PldnewVb"
ElseIf optF.Value = True Then
    tt = "Select * from PldnewVc"
ElseIf optG.Value = True Then
    tt = "Select * from PldnewVad"
ElseIf optH.Value = True Then
    tt = "Select * from PldnewVb"
ElseIf optI.Value = True Then
    tt = "Select * from PldnewVc"
ElseIf optJ.Value = True Then
    tt = "Select * from PldnewVae"
ElseIf optK.Value = True Then
    tt = "Select * from PldnewVb"
ElseIf optL.Value = True Then
    tt = "Select * from PldnewVc"
ElseIf optM.Value = True Then
    tt = "Select * from PldnewVab"
ElseIf optN.Value = True Then
    tt = "Select * from PldnewVb"
ElseIf optO.Value = True Then
    tt = "Select * from PldnewVc"
ElseIf optP.Value = True Then
    tt = "Select * from PldnewVa"
ElseIf optQ.Value = True Then
    tt = "Select * from PldnewVb"
ElseIf optR.Value = True Then
    tt = "Select * from PldnewVc"
ElseIf optZa.Value = True Or optZb.Value = True Or optZc.Value = True Or _
       optZd.Value = True Or optZe.Value = True Or optZf.Value = True Then                         '作废单
    tt = "Select * from PldZfV"
End If
    mod1.PldV.Close
    mod1.PldV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set dtgPld.DataSource = mod1.PldV
End Sub

Private Sub Form_Load()
frmA.Left = 6120
frmA.Top = 1260
frmB.Left = 6120
frmB.Top = 1260
frmC.Left = 6120
frmC.Top = 1260
frmD.Left = 6120
frmD.Top = 1260
frmE.Left = 6120
frmE.Top = 1260
frmF.Left = 6120
frmF.Top = 1260
frmA.BorderStyle = 0
frmB.BorderStyle = 0
frmC.BorderStyle = 0
frmD.BorderStyle = 0
frmE.BorderStyle = 0
frmF.BorderStyle = 0
frmPVa.Height = 4770
frmPVa.Width = 7575
opTa.Value = False
opTB.Value = False
opTC.Value = False
optD.Value = False
optE.Value = False
optF.Value = False
optG.Value = False
optH.Value = False
optI.Value = False
optJ.Value = False
optK.Value = False
optL.Value = False
optM.Value = False
optN.Value = False
optO.Value = False
optP.Value = False
optR.Value = False
optZa.Value = False
optZb.Value = False
optZc.Value = False
optZd.Value = False
optZe.Value = False
optZf.Value = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmPVa.Visible = False
frmZu.Enabled = True
Cancel = True
End Sub

Private Sub txtBr_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tt As String
On Error Resume Next
If KeyCode = 13 Then
    Select Case comBR.Text
    Case "合同金额"
        'tt = "PLDV_htze(" & Val(txtBr.Text) & ")"
        tt = "Select * from PLDV where htze=" & Val(txtBr.Text)
    Case "合同编号"
        'tt = "PLDV_htbh('" & txtBr.Text & "')"
        tt = "Select * from PLDV where htbh='" & txtBr.Text & "'"
    Case "项目名称"
        'tt = "PLDV_xmmc('" & txtBr.Text & "')"
        tt = "Select * from PLDV where xmmc like '%" & txtBr.Text & "%'"
    End Select
    mod1.PldV.Close
    mod1.PldV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set dtgPld.DataSource = mod1.PldV
End If
End Sub
