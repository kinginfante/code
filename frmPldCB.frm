VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmPldCB 
   Caption         =   "成本结算"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   12660
   Begin MSAdodcLib.Adodc adoDy 
      Height          =   330
      Left            =   11280
      Top             =   2340
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc adoHp 
      Height          =   330
      Left            =   11280
      Top             =   1680
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.CommandButton cmdJz 
      Caption         =   "总结账"
      Height          =   435
      Left            =   11430
      TabIndex        =   4
      Top             =   180
      Width           =   945
   End
   Begin VB.TextBox txtBz 
      Height          =   2205
      Left            =   10230
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3780
      Width           =   2385
   End
   Begin MSDataGridLib.DataGrid dtgHp 
      Bindings        =   "frmPldCB.frx":0000
      Height          =   3465
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   6112
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
      TabAction       =   2
      WrapCellPointer =   -1  'True
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
         DataField       =   "ljMc"
         Caption         =   "产品名称"
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
         DataField       =   "phBiao"
         Caption         =   "牌号商标"
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
         DataField       =   "ljBh"
         Caption         =   "规格型号"
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
         DataField       =   "jlDw"
         Caption         =   "单位"
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
         DataField       =   "KzSl"
         Caption         =   "库存数量"
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
         DataField       =   "kcdj"
         Caption         =   "库存单价"
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
         DataField       =   "CzSL"
         Caption         =   "采购数量"
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
         DataField       =   "CGDJ"
         Caption         =   "采购单价"
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
         DataField       =   "JZF"
         Caption         =   "结帐否"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "已结帐"
            FalseValue      =   "未结"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   7
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   2399.811
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   794.835
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dtgDy 
      Bindings        =   "frmPldCB.frx":0014
      Height          =   2235
      Left            =   0
      TabIndex        =   1
      Top             =   3780
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   3942
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
      TabAction       =   2
      WrapCellPointer =   -1  'True
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
      Caption         =   "多余采购"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "ljMc"
         Caption         =   "产品名称"
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
         DataField       =   "phBiao"
         Caption         =   "牌号商标"
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
         DataField       =   "ljBh"
         Caption         =   "规格型号"
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
         DataField       =   "jlDw"
         Caption         =   "单位"
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
         DataField       =   "SL"
         Caption         =   "数量"
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
         DataField       =   "dj"
         Caption         =   "单价"
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
         DataField       =   "je"
         Caption         =   "金额"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "JZF"
         Caption         =   "结帐否"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "√"
            FalseValue      =   ""
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   7
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         BeginProperty Column00 
            Locked          =   -1  'True
            Object.Visible         =   -1  'True
            ColumnWidth     =   2399.811
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            ColumnWidth     =   2099.906
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column06 
            Locked          =   -1  'True
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column07 
            Button          =   -1  'True
            ColumnWidth     =   705.26
         EndProperty
      EndProperty
   End
   Begin VB.Label lblJe 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11340
      TabIndex        =   5
      Top             =   900
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "多购原因:"
      Height          =   225
      Left            =   11280
      TabIndex        =   3
      Top             =   3420
      Width           =   1335
   End
End
Attribute VB_Name = "frmPldCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdJZ_Click()
Dim tt As String
On Error Resume Next
Dim Tze As Single
Dim Thp As Single
Dim Tdy As Single
Tze = 0
Thp = 0
Tdy = 0

If frmPld.lblZT.Caption = "此单已经作废" Then Exit Sub
'If mod1.PLE = True And frmPld.cmdQME.Caption = "" Then
If frmPld.lblLc.Caption = 5 Then
    '非维保合同,一次性结帐
    If frmPld.lblXZ <> "WB" And frmPld.lblXZ <> "WX" Then
        ii = MsgBox("此配料单任务已经全部完成,您确认结算成本吗?", vbInformation + vbYesNo, "Hello!")
        If ii = vbYes Then
            frmPld.lblLc.Caption = 6
            frmPld.cmdQME.Caption = mod1.DName
            frmPld.lblTe.Caption = mod1.DQda
    
    
            adoHp.Recordset.MoveFirst
            Do While Not adoHp.Recordset.EOF
                Thp = Thp + adoHp.Recordset.Fields("je").Value
                adoHp.Recordset.MoveNext
            Loop
            
            adoDy.Recordset.MoveFirst
            Do While Not adoDy.Recordset.EOF
                If adoDy.Recordset.Fields("jzF").Value = True Then
                    Tdy = Tdy + adoDy.Recordset.Fields("je").Value
                End If
                adoDy.Recordset.MoveNext
            Loop
            Tze = Thp + Tdy
            lblJe.Caption = Tze
            frmPld.txtCB.Text = Tze
            
            '更新合同评审单的其它成本和提成
            tt = "Htcb_clcb('" & frmPld.txtHtbh.Text & "'," & Tze & ")"
            Set mod1.HTP = New ADODB.Recordset
            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
            
            '签字保存
            adoHp.Recordset.UpdateBatch
            adoDy.Recordset.UpdateBatch
            
            tt = "PLDBoundA(" & frmPld.lblPmid.Caption & ")"
            Set mod1.HTP = New ADODB.Recordset
            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
            
            mod1.HTP.Update "LC", frmPld.lblLc.Caption '流程
            mod1.HTP.Update "QME", frmPld.cmdQME.Caption
            mod1.HTP.Update "QMET", frmPld.lblTe.Caption
            mod1.HTP.Update "BZE", frmPld.lblBe.Caption '开单备注
            mod1.HTP.Update "Tze", frmPld.txtCB.Text '成本总额
            mod1.HTP.Update "pwf", 1
            mod1.HTP.UpdateBatch
            
            Call modPld.PldJl(frmPld.lblPmid.Caption)
            
            '更新浏览列表
            If frmHtZX.Visible = True Then
                frmHtZX.adoPld.Requery
                Set frmHtZX.dtgPld.DataSource = frmHtZX.adoPld
            ElseIf frmHtZxG.Visible = True Then
                frmHtZxG.adoPld.Requery
                Set frmHtZxG.dtgPld.DataSource = frmHtZxG.adoPld
            ElseIf Dialog.Visible = True Then
                Call mod1.refEnvent
            End If
            frmPld.cmdSave.Enabled = False
            Call mod1.EnventFinish(frmPld.lblFwid.Caption)
'            tt = "update pldMain set Pwf=1 where pmid=" & Val(lblPmid.Caption)
'            Set mod1.HTP = New ADODB.Recordset
'            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
        End If
    Else '维保维修合同,非一次性结帐
        ii = MsgBox("此项维保配料任务已经完成,您确认结算这笔成本吗?", vbInformation + vbYesNo, "Hello!")
        If ii = vbYes Then
            frmPld.lblLc.Caption = 6
            frmPld.cmdQME.Caption = mod1.DName
            frmPld.lblTe.Caption = mod1.DQda
    
    
            adoHp.Recordset.MoveFirst
            Do While Not adoHp.Recordset.EOF
                If adoHp.Recordset.Fields("jzF").Value = 0 Then
                Thp = Thp + adoHp.Recordset.Fields("je").Value
                adoHp.Recordset.Fields("jzF").Value = 1
                End If
                adoHp.Recordset.MoveNext
            Loop
            
            adoDy.Recordset.MoveFirst
            Do While Not adoDy.Recordset.EOF
                If adoDy.Recordset.Fields("jzF").Value = True Then
                    Tdy = Tdy + adoDy.Recordset.Fields("je").Value
                End If
                adoDy.Recordset.MoveNext
            Loop
            Tze = Thp + Tdy
            lblJe.Caption = lblJe.Caption + Tze
            frmPld.txtCB.Text = lblJe.Caption
            
            '更新合同评审单的其它成本和提成
            tt = "Htcb_clcb('" & frmPld.txtHtbh.Text & "'," & Tze & ")"
            Set mod1.HTP = New ADODB.Recordset
            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
            
            '签字保存
            adoHp.Recordset.UpdateBatch
            adoDy.Recordset.UpdateBatch
            
            tt = "PLDBoundA(" & frmPld.lblPmid.Caption & ")"
            Set mod1.HTP = New ADODB.Recordset
            mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdStoredProc
            
            mod1.HTP.Update "LC", frmPld.lblLc.Caption '流程
            mod1.HTP.Update "QME", frmPld.cmdQME.Caption
            mod1.HTP.Update "QMET", frmPld.lblTe.Caption
            mod1.HTP.Update "BZE", frmPld.lblBe.Caption '开单备注
            mod1.HTP.Update "Tze", frmPld.txtCB.Text '成本总额
            mod1.HTP.Update "pwf", 1
            mod1.HTP.UpdateBatch
            
            Call modPld.PldJl(frmPld.lblPmid.Caption)
            
            '更新浏览列表
            If frmHtZX.Visible = True Then
                frmHtZX.adoPld.Requery
                Set frmHtZX.dtgPld.DataSource = frmHtZX.adoPld
            ElseIf frmHtZxG.Visible = True Then
                frmHtZxG.adoPld.Requery
                Set frmHtZxG.dtgPld.DataSource = frmHtZxG.adoPld
            ElseIf Dialog.Visible = True Then
                Call mod1.refEnvent
            End If
            Call mod1.EnventFinish(frmPld.lblFwid.Caption)
            frmPld.cmdSave.Enabled = False
        End If
    End If
End If
End Sub



















Private Sub dtgDy_AfterColUpdate(ByVal ColIndex As Integer)
On Error Resume Next
adoDy.Recordset.Fields("je").Value = adoDy.Recordset.Fields("sl").Value * adoDy.Recordset.Fields("dj").Value
End Sub

Private Sub dtgDy_ButtonClick(ByVal ColIndex As Integer)
If ColIndex = 7 Then
    If adoDy.Recordset.Fields("jzF").Value = True Then
        adoDy.Recordset.Fields("jzF").Value = False
    Else
        adoDy.Recordset.Fields("jzF").Value = True
    End If
    Set dtgDy.DataSource = adoDy.Recordset
End If
End Sub

Private Sub dtgDy_Click()
On Error Resume Next
txtBz.Text = adoDy.Recordset.Fields("Bz").Value
End Sub

Private Sub dtgDy_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
txtBz.Text = adoDy.Recordset.Fields("Bz").Value
End Sub


Private Sub dtgHp_AfterColUpdate(ByVal ColIndex As Integer)
On Error Resume Next
adoHp.Recordset.Fields("je").Value = adoHp.Recordset.Fields("kzsl").Value * adoHp.Recordset.Fields("kcdj").Value + _
        adoHp.Recordset.Fields("czsl").Value * adoHp.Recordset.Fields("cgdj").Value
End Sub

Private Sub Form_Load()
frmPldCB.Width = 12780
frmPldCB.Height = 6435
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmPldCB.Visible = False
Cancel = True
frmPld.Enabled = True
frmPld.ZOrder 0
End Sub

Private Sub txtBz_LostFocus()
adoDy.Recordset.Update "Bz", txtBz.Text
End Sub


