VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmFuK 
   Caption         =   "付款方式表"
   ClientHeight    =   4185
   ClientLeft      =   240
   ClientTop       =   3315
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   8190
   Begin VB.CommandButton cmdAdd 
      Caption         =   "添加"
      Height          =   375
      Left            =   390
      TabIndex        =   9
      Top             =   1620
      Width           =   525
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除"
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   1620
      Width           =   495
   End
   Begin VB.CommandButton cmdMod1 
      Caption         =   "修改"
      Height          =   375
      Left            =   1500
      TabIndex        =   7
      Top             =   1620
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdMod2 
      Caption         =   "修改"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3810
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdYdel 
      Caption         =   "删除"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1020
      TabIndex        =   3
      Top             =   3810
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdYadd 
      Caption         =   "添加"
      Enabled         =   0   'False
      Height          =   375
      Left            =   450
      TabIndex        =   2
      Top             =   3810
      Visible         =   0   'False
      Width           =   525
   End
   Begin MSDataGridLib.DataGrid dtgYf 
      Bindings        =   "frmFuK.frx":0000
      Height          =   1725
      Left            =   0
      TabIndex        =   1
      Top             =   2010
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3043
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   13631199
      ForeColor       =   12582912
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   18
      BeginProperty Column00 
         DataField       =   "YiRq"
         Caption         =   "收款日期"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dddddd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "yfJe"
         Caption         =   "收款金额"
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
      BeginProperty Column03 
         DataField       =   "yingRQ"
         Caption         =   "yingRQ"
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
         DataField       =   "htF"
         Caption         =   "htF"
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
         DataField       =   "zcF"
         Caption         =   "zcF"
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
         DataField       =   "yWy"
         Caption         =   "yWy"
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
         DataField       =   "YingJe"
         Caption         =   "YingJe"
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
         DataField       =   "DelF"
         Caption         =   "DelF"
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
         DataField       =   "khMc"
         Caption         =   "khMc"
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
         DataField       =   "fkFc"
         Caption         =   "付款方式"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "银行"
            FalseValue      =   "现金"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "yinHang"
         Caption         =   "银 行"
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
         DataField       =   "qianKuan1"
         Caption         =   "qianKuan1"
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
      BeginProperty Column13 
         DataField       =   "qianKuan2"
         Caption         =   "qianKuan2"
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
      BeginProperty Column14 
         DataField       =   "qianKuan3"
         Caption         =   "qianKuan3"
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
      BeginProperty Column15 
         DataField       =   "qianKuan4"
         Caption         =   "qianKuan4"
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
      BeginProperty Column16 
         DataField       =   "qianKuan5"
         Caption         =   "qianKuan5"
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
      BeginProperty Column17 
         DataField       =   "qianKuan6"
         Caption         =   "qianKuan6"
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
         ScrollBars      =   2
         BeginProperty Column00 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814.74
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column10 
            Button          =   -1  'True
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoYf 
      Height          =   375
      Left            =   2130
      Top             =   3060
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   661
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
   Begin VB.CommandButton cmdGx 
      Caption         =   "关 闭"
      Height          =   345
      Left            =   7530
      TabIndex        =   0
      Top             =   3810
      Width           =   645
   End
   Begin MSAdodcLib.Adodc adoHpt 
      Height          =   405
      Left            =   1740
      Top             =   2820
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin MSDataGridLib.DataGrid dtgFk 
      Bindings        =   "frmFuK.frx":0014
      Height          =   2025
      Left            =   0
      TabIndex        =   10
      Top             =   -30
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   3572
      _Version        =   393216
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "rq"
         Caption         =   "应付日期"
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
         DataField       =   "ED"
         Caption         =   "收款额度"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   5
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "yingfJe"
         Caption         =   "应收金额"
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
         DataField       =   "ZT"
         Caption         =   "状态"
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
         DataField       =   "zcF"
         Caption         =   "付款否"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "OK!"
            FalseValue      =   "欠款"
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
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3300.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   705.26
         EndProperty
      EndProperty
   End
   Begin VB.Label Label38 
      Caption         =   "合同总额："
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   2070
      Width           =   945
   End
   Begin VB.Label lblHtZe 
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   2400
      Width           =   1245
   End
End
Attribute VB_Name = "frmFuK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim yT As String
Dim YiDate As Date ' 收款表日期，用于监测用户日期输入格式是否正确
Dim addB As Boolean '是否添加应收款记录
Dim OldRq As Date '旧的日期，用来 修改应收表时，相应的资金流量表应收总额改变
Dim OldJe As Single '旧的应收金额，用来 修改应收表时，相应的资金流量表应收总额改变

Private Sub cmdAdd_Click()
If form2Htp.optW.Value = True Then Exit Sub

frmFuK.adoHpt.Recordset.AddNew "htbh", form2Htp.txtHtbh.Text
If form2Htp.optP.Value = True Then
frmFuK.adoHpt.Recordset.Update "htF", 0
ElseIf form2Htp.optZ.Value = True Then
frmFuK.adoHpt.Recordset.Update "htF", 1
End If
frmFuK.adoHpt.Recordset.Update "delF", 1
frmFuK.adoHpt.Recordset.Update "zcF", 0
frmFuK.adoHpt.Recordset.Update "khMc", form2Htp.txtKhmc.Text
frmFuK.adoHpt.Recordset.Update "yWy", form2Htp.txtYwy.Text
frmFuK.adoHpt.Recordset.Update "yifJe", 0
Set dtgFk.DataSource = frmFuK.adoHpt
End Sub

Private Sub Command1_Click()
dtgYf.Columns(1).Visible = False
End Sub

Private Sub cmdMod1_Click()
If adoHpt.Recordset.RecordCount = 0 Then
    frmFuK.adoHpt.Recordset.AddNew "htbh", form2Htp.txtHtbh.Text
    If form2Htp.optP.Value = True Then
    frmFuK.adoHpt.Recordset.Update "htF", 0
    ElseIf form2Htp.optZ.Value = True Then
    frmFuK.adoHpt.Recordset.Update "htF", 1
    End If
    frmFuK.adoHpt.Recordset.Update "delF", 1
    frmFuK.adoHpt.Recordset.Update "zcF", 0
    frmFuK.adoHpt.Recordset.Update "khMc", form2Htp.txtKhmc.Text
    frmFuK.adoHpt.Recordset.Update "yWy", form2Htp.txtYwy.Text
    frmFuK.adoHpt.Recordset.Update "yifJe", 0
    Set dtgFk.DataSource = frmFuK.adoHpt
End If
cmdDel.Enabled = True
cmdAdd.Enabled = True
dtgFk.AllowUpdate = True
cmdMod1.Enabled = False
End Sub

Private Sub cmdMod2_Click()
cmdYadd.Enabled = True
cmdYdel.Enabled = True
dtgYf.AllowUpdate = True
dtgYf.Enabled = True
cmdMod2.Enabled = False
End Sub

Private Sub cmdYadd_Click()
On Error Resume Next
adoYf.Recordset.AddNew "yfJe", 0
adoYf.Recordset.Update "fkfc", 0
adoYf.Recordset.Update "qiankuan1", 1
adoYf.Recordset.Update "qiankuan2", 1
adoYf.Recordset.Update "qiankuan3", 1
adoYf.Recordset.Update "qiankuan4", 1
adoYf.Recordset.Update "qiankuan5", 1
adoYf.Recordset.Update "qiankuan6", 1
adoYf.Recordset.Update "delF", 1
adoYf.Recordset.Update "htF", 1
adoYf.Recordset.UpdateBatch
Set dtgYf.DataSource = adoYf
End Sub

Private Sub dtgFk_AfterColEdit(ByVal ColIndex As Integer)
If ColIndex = 1 Then
frmFuK.adoHpt.Recordset.Update "ED", frmFuK.adoHpt.Recordset.Fields("ED").Value / 100
frmFuK.adoHpt.Recordset.Update "yingfJe", Val(lblHtze.Caption) * frmFuK.adoHpt.Recordset.Fields("ED").Value
ElseIf ColIndex = 2 Then
frmFuK.adoHpt.Recordset.Update "ED", frmFuK.adoHpt.Recordset.Fields("yingfJe").Value / Val(lblHtze.Caption)
End If
'Set dtgFk.DataSource = frmFuK.adoHpt

End Sub

Private Sub dtgFk_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
'On Error Resume Next
'YiDate = dtgFk.Columns(1).Value
'OldRq = dtgFk.Columns(1).Value
'OldJe = 0
'If ColIndex = 1 Then
'OldRq = OldValue
'If OldValue <> "" Then
'Cancel = True
'End If
'ElseIf ColIndex = 2 Then
'OldJe = OldValue
'End If


End Sub

Private Sub dtgFk_BeforeInsert(Cancel As Integer)
'frmFuK.adoHpt.Recordset.Fields(0).Value = form2Htp.txtYwy.Text
'frmFuK.adoHpt.Recordset.Fields(3).Value = form2Htp.txtHtbh.Text
'frmFuK.adoHpt.Recordset.UpdateBatch
End Sub

Private Sub dtgFk_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
If KeyCode = 13 Then SendKeys "{tab}"
'frmFuK.adoHpt.Recordset.Update "ywy", form2Htp.txtYwy.Text
'frmFuK.adoHpt.Recordset.Update "htBh", form2Htp.txtHtbh.Text
'frmFuK.adoHpt.Recordset.Update "khmc", form2Htp.txtKhmc.Text
'frmFuK.adoHpt.Recordset.Update "delF", 1  '删除否，合同未删除
'frmFuK.adoHpt.Recordset.Update "zcF", 0  '收到否，由于为评审阶段，所以未收到
'frmFuK.adoHpt.Recordset.Update "htF", 0 '合同执行否，由于为评审阶段，所以未执行

End Sub

Private Sub dtgYf_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
If ColIndex = 0 Then
'如果键入行为当前应付日期（与应收日期一致的日期），则此行日期不能修改
On Error GoTo yFErr
'dtgFk.Columns(1).Value = Format(dtgFk.Columns(1).Value, "long date")
OldValue = Format(OldValue, "YYYY-M-D")
If OldValue = dtgFk.Columns(1).Value Then
Cancel = True
dtgYf.Columns(0).Value = OldValue
Exit Sub
End If
YiDate = dtgYf.Columns(0).Value
OldRq = OldValue
End If



On Error Resume Next
'        '先得出当前日期流量表中的以收金额总数
'        Dim jT As String
'        jT = "Select zJ from llb1 where rq='" & adoYf.Recordset.Fields(0).Value & "'"
'        mod1.zjJinT.Close
'        mod1.zjJinT.Open jT, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
'        If IsNull(mod1.zjJinT.Fields(0).Value) = True Then mod1.zjJinT.Fields(0).Value = 0
'        '再减去当前的旧的已收金额，使得更新后可以加上新的已收金额
'        mod1.zjJinT.Fields(0).Value = mod1.zjJinT.Fields(0).Value - OldValue

Exit Sub

yFErr:
'YiDate = Null
'dtgYf.Columns(0).Value = Null
End Sub

Private Sub dtgYf_ButtonClick(ByVal ColIndex As Integer)
If adoYf.Recordset.Fields(10).Value = 0 Then
adoYf.Recordset.Update "fkfc", 1

Else
adoYf.Recordset.Update "fkfc", 0
End If
Set dtgYf.DataSource = adoYf
End Sub

Private Sub dtgYf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
'If frmFuK.dtgYf.Columns(0).Value = "" Or frmFuK.dtgYf.Columns(1).Value = "" Then
'frmFuK.dtgYf.Columns(0).Value = "": frmFuK.dtgYf.Columns(1).Value = ""
'Cancel = True
'Else
On Error Resume Next
frmFuK.adoYf.Recordset.Update "htBh", frmFuK.adoHpt.Recordset.Fields(3).Value
frmFuK.adoYf.Recordset.Update "yingRQ", frmFuK.adoHpt.Recordset.Fields(1).Value
frmFuK.adoYf.Recordset.Update "ywy", frmFuK.adoHpt.Recordset.Fields(0).Value
frmFuK.adoYf.Recordset.Update "zcF", frmFuK.adoHpt.Recordset(5).Value
frmFuK.adoYf.Recordset.Update "YingJe", frmFuK.adoHpt.Recordset(2).Value
frmFuK.adoYf.Recordset.Update "delF", 1  '删除否，合同未删除
'frmFuK.adoYf.Recordset.Update "zcF", 0  '收到否，由于为评审阶段，所以未收到
frmFuK.adoYf.Recordset.Update "htF", 1 '合同执行否，由于为评审阶段，所以未执行
frmFuK.adoYf.Recordset.Update "khmc", frmFuK.adoHpt.Recordset(8).Value
'End If
End Sub



Private Sub Form_Load()
frmFuK.Width = 8340
frmFuK.Height = 4620
frmFuK.Left = 2000
frmFuK.Top = 3000
addB = False
End Sub

