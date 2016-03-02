VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmWbBj 
   Caption         =   "维保清单"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   12885
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4290
   ScaleWidth      =   12885
   Begin MSDataListLib.DataCombo comLb 
      Height          =   330
      Left            =   5310
      TabIndex        =   6
      Top             =   30
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid dtGNr 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   1050
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   5741
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   14
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
         DataField       =   "wNr"
         Caption         =   "服务内容"
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
         DataField       =   "wbx"
         Caption         =   "形式"
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
         DataField       =   "KXF"
         Caption         =   "基本选项"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   ""
            FalseValue      =   "√"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "gT"
         Caption         =   "工时"
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
         DataField       =   "dGt"
         Caption         =   "附加工时"
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
         DataField       =   "dw"
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
      BeginProperty Column06 
         DataField       =   "fjL"
         Caption         =   "附加计算类型"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "按附加量计算"
            FalseValue      =   "按机器数量计算"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "BF"
         Caption         =   "设备费"
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
         DataField       =   "BZ"
         Caption         =   "备注"
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
         ScrollBars      =   3
         RecordSelectors =   0   'False
         Size            =   457
         BeginProperty Column00 
            Object.Visible         =   -1  'True
            ColumnWidth     =   3600
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   -1  'True
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   -1  'True
            ColumnWidth     =   2294.929
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo comXh 
      Height          =   330
      Left            =   1560
      TabIndex        =   1
      Top             =   570
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo comPb 
      Height          =   330
      Left            =   1560
      TabIndex        =   2
      Top             =   0
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   582
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.Label Label4 
      Caption         =   "双击列表可将该项添加进服务表"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10230
      TabIndex        =   7
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "系统类别:"
      Height          =   285
      Left            =   4140
      TabIndex        =   5
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "机组型号:"
      Height          =   285
      Left            =   270
      TabIndex        =   4
      Top             =   630
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "机组品牌:"
      Height          =   345
      Left            =   270
      TabIndex        =   3
      Top             =   30
      Width           =   1125
   End
End
Attribute VB_Name = "frmWbBj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoBj As ADODB.Recordset
Public adoTl As ADODB.Recordset


Private Sub comLb_Click(Area As Integer)

Dim tt As String
Dim wbx As String

If frmWbBj.Visible = False Or comLb.Text = "" Then Exit Sub
On Error Resume Next
If frmWBXJ.tabGc.Tab = 0 Then
    wbx = "年保"
ElseIf frmWBXJ.tabGc.Tab = 1 Then
    wbx = "例检"
End If
tt = "select * from bjxtview where xtid='" & comLb.BoundText & "' and wbx='" & wbx & "'"
adoTl.Close
adoTl.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtGNr.DataSource = adoTl
End Sub

Private Sub dtGNr_DblClick()
Dim tt As String
Dim Dl As String
Dim Kxf As String
On Error Resume Next
Dl = ""
Kxf = ""
If adoTl.Fields("fjl").Value = True Then
    Dl = InputBox("请输入附加量!", "询问")
    If Val(Dl) = 0 Then
        Exit Sub
    End If
End If
Set mod1.CMD = New ADODB.command
mod1.CMD.ActiveConnection = mod1.CC
mod1.CMD.CommandText = "xunJiaWbAdd"
mod1.CMD.CommandType = adCmdStoredProc
mod1.CMD.Parameters("@jzPb") = adoTl.Fields("jzpb").Value
mod1.CMD.Parameters("@jzXh") = adoTl.Fields("jzXh").Value
mod1.CMD.Parameters("@XT") = adoTl.Fields("XT").Value
If adoTl.Fields("kxf").Value = True Then
    Kxf = "可选"
End If
mod1.CMD.Parameters("@kxf") = Kxf
mod1.CMD.Parameters("@wbX") = adoTl.Fields("wbX").Value
mod1.CMD.Parameters("@wNr") = adoTl.Fields("wNr").Value
mod1.CMD.Parameters("@gT") = adoTl.Fields("gT").Value
mod1.CMD.Parameters("@dGt") = adoTl.Fields("dGt").Value
mod1.CMD.Parameters("@dw") = adoTl.Fields("dw").Value
mod1.CMD.Parameters("@fjL") = adoTl.Fields("fjL").Value
mod1.CMD.Parameters("@BF") = adoTl.Fields("BF").Value
mod1.CMD.Parameters("@BZ") = adoTl.Fields("BZ").Value
mod1.CMD.Parameters("@wId") = adoTl.Fields("wId").Value
mod1.CMD.Parameters("@xtId") = adoTl.Fields("xtId").Value
mod1.CMD.Parameters("@bid") = frmWBXJ.lblBid.Caption
mod1.CMD.Parameters("@dl") = Val(Dl)                         '附加单位量
mod1.CMD.Parameters("@jCount") = Val(frmWBXJ.txtSl.Text)
mod1.CMD.Execute
Set CMD = Nothing

If frmWBXJ.tabGc.Tab = 0 Then '刷新年保表
tt = "select * from xunJIaWbView where wbx='年保' and bid=" & Val(frmWBXJ.lblBid.Caption)
    frmWBXJ.adoWb.Close
    frmWBXJ.adoWb.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWBXJ.dtgWb.DataSource = frmWBXJ.adoWb
    'If frmWBXJ.adoWb.RecordCount > 1 Then
        frmWBXJ.dtgWb.FixedRows = 0
        frmWBXJ.dtgWb.MergeCol(1) = True
        frmWBXJ.dtgWb.MergeCol(2) = True
        frmWBXJ.dtgWb.MergeCol(3) = True
        frmWBXJ.dtgWb.MergeCells = 3
        frmWBXJ.dtgWb.FixedRows = 1
    'End If
ElseIf frmWBXJ.tabGc.Tab = 1 Then                              '刷新例检表
tt = "select * from xunJIaWbView where wbx='例检' and bid=" & Val(frmWBXJ.lblBid.Caption)
    frmWBXJ.adoLj.Close
    frmWBXJ.adoLj.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
    Set frmWBXJ.dtgLj.DataSource = frmWBXJ.adoLj
    'If frmWBXJ.adoLj.RecordCount > 1 Then
        frmWBXJ.dtgLj.FixedRows = 0
        frmWBXJ.dtgLj.MergeCol(1) = True
        frmWBXJ.dtgLj.MergeCol(2) = True
        frmWBXJ.dtgLj.MergeCol(3) = True
        frmWBXJ.dtgLj.MergeCells = 3
        frmWBXJ.dtgLj.FixedRows = 1
    'End If
End If

End Sub

Private Sub Form_Load()
Set adoBj = New ADODB.Recordset
Set adoTl = New ADODB.Recordset
frmWbBj.Height = 4860
frmWbBj.Width = 13005
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
frmWbBj.Visible = False
frmWBXJ.Enabled = True
frmWBXJ.ZOrder 0
Cancel = True
End If
End Sub
