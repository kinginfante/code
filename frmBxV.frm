VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form frmBxV 
   Caption         =   "报销单查询"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdBack 
      Caption         =   "导航"
      Height          =   585
      Left            =   14400
      Picture         =   "frmBxV.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8370
      Width           =   675
   End
   Begin VB.Frame frmBxVM 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   15195
      Begin VB.CommandButton cmdBr 
         Caption         =   "..."
         Height          =   255
         Left            =   3960
         TabIndex        =   30
         Top             =   6330
         Width           =   405
      End
      Begin NiceFormControl.NiceButton NiceButton1 
         Height          =   315
         Left            =   300
         TabIndex        =   29
         Top             =   8640
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         BTYPE           =   1
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   12648384
         BCOLO           =   16777152
         FCOL            =   0
         FCOLO           =   12640511
         MCOL            =   12648384
         MPTR            =   1
         MICON           =   "frmBxV.frx":0102
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
         Style           =   20
         Caption         =   "添加保存"
      End
      Begin VB.Timer timWait 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1350
         Top             =   420
      End
      Begin VB.Timer timQuit 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1770
         Top             =   0
      End
      Begin VB.TextBox Text3 
         Height          =   705
         Left            =   1530
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   7770
         Width           =   2865
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   1530
         TabIndex        =   25
         Top             =   6810
         Width           =   2865
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1530
         TabIndex        =   24
         Top             =   6330
         Width           =   2325
      End
      Begin MSComCtl2.DTPicker dtpBrq 
         Height          =   315
         Left            =   1530
         TabIndex        =   23
         Top             =   5760
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   8454016
         CalendarTitleBackColor=   16711808
         CalendarTrailingForeColor=   -2147483635
         Format          =   101056513
         CurrentDate     =   38797
      End
      Begin MSComCtl2.DTPicker dtpDZ 
         Height          =   315
         Left            =   1530
         TabIndex        =   26
         Top             =   7260
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   8454016
         CalendarTitleBackColor=   16711808
         CalendarTrailingForeColor=   -2147483635
         Format          =   101056513
         CurrentDate     =   38797
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgDZ 
         Height          =   5445
         Left            =   0
         TabIndex        =   28
         Top             =   30
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   9604
         _Version        =   393216
         BackColor       =   16777152
         Rows            =   8
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   15728356
         BackColorBkg    =   16777152
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   3
         PictureType     =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         Height          =   255
         Left            =   330
         TabIndex        =   22
         Top             =   7890
         Width           =   945
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "到帐日期"
         Height          =   255
         Left            =   330
         TabIndex        =   21
         Top             =   7395
         Width           =   1035
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "金额"
         Height          =   255
         Left            =   330
         TabIndex        =   20
         Top             =   6870
         Width           =   795
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "客户名称"
         Height          =   255
         Left            =   330
         TabIndex        =   19
         Top             =   6345
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "公告日期"
         Height          =   255
         Left            =   330
         TabIndex        =   18
         Top             =   5820
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCWBB 
      BackColor       =   &H00C0FFC0&
      Caption         =   "2008新财务报表"
      Height          =   375
      Left            =   10950
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7140
      Width           =   2265
   End
   Begin VB.CommandButton cmdV 
      Caption         =   "查  询"
      Height          =   285
      Left            =   12270
      TabIndex        =   15
      Top             =   8550
      Width           =   885
   End
   Begin VB.TextBox txtZ 
      Height          =   270
      Left            =   10830
      TabIndex        =   14
      Top             =   8550
      Width           =   1305
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "frmBxV.frx":011E
      Left            =   9450
      List            =   "frmBxV.frx":012E
      TabIndex        =   12
      Text            =   "金额"
      Top             =   8550
      Width           =   945
   End
   Begin VB.CommandButton cmdNB 
      Caption         =   "内部结算报销单"
      Height          =   345
      Left            =   10950
      TabIndex        =   10
      Top             =   4710
      Width           =   2265
   End
   Begin VB.CommandButton cmdFybb 
      Caption         =   "费用报表"
      Height          =   375
      Left            =   10950
      TabIndex        =   9
      Top             =   6420
      Width           =   2295
   End
   Begin VB.CommandButton cmdWB 
      Caption         =   "未 报 销 单 据"
      Height          =   345
      Left            =   10920
      TabIndex        =   8
      Top             =   5220
      Width           =   2325
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "详  情"
      Height          =   345
      Left            =   10980
      TabIndex        =   7
      Top             =   180
      Width           =   4005
   End
   Begin VB.CommandButton cmdFw 
      Caption         =   "查询范围"
      Height          =   315
      Left            =   11010
      TabIndex        =   5
      Top             =   3750
      Width           =   1095
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "上月"
      Height          =   345
      Left            =   13290
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   "下月"
      Height          =   345
      Left            =   14190
      TabIndex        =   2
      Top             =   2880
      Width           =   825
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBx 
      Height          =   8985
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   15849
      _Version        =   393216
      BackColorBkg    =   -2147483636
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.MonthView mtA 
      Height          =   2160
      Left            =   10980
      TabIndex        =   4
      Top             =   630
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   3810
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      MonthBackColor  =   -2147483633
      ShowToday       =   0   'False
      StartOfWeek     =   101056513
      TitleBackColor  =   16711935
      CurrentDate     =   38666
   End
   Begin VB.Label Label2 
      Caption         =   "值"
      Height          =   225
      Left            =   10500
      TabIndex        =   13
      Top             =   8580
      Width           =   225
   End
   Begin VB.Label Label1 
      Caption         =   "类别"
      Height          =   255
      Left            =   9000
      TabIndex        =   11
      Top             =   8580
      Width           =   405
   End
   Begin VB.Line Line2 
      X1              =   8850
      X2              =   15240
      Y1              =   6180
      Y2              =   6180
   End
   Begin VB.Line Line1 
      X1              =   8970
      X2              =   15240
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label lblFw 
      Height          =   285
      Left            =   12210
      TabIndex        =   6
      Top             =   3780
      Width           =   1035
   End
End
Attribute VB_Name = "frmBxV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoBxV As ADODB.Recordset
Dim timZm As Integer '数据提交后,由timWait执行的后续命令ID(1到帐添加,5报销添加)

Private Sub cmdBack_Click()
frmBxV.Visible = False
frmZu.Enabled = True

End Sub

Private Sub cmdCWBB_Click()
Me.Enabled = False
frmCWBB.Show
frmCWBB.ZOrder 0
End Sub

Private Sub cmdFw_Click()
'frmBxV.Enabled = False
Set Ren.XForm = New frmBxV
Call mod1.RenXz("frmBxV", Me, 0)

End Sub

Private Sub cmdFybb_Click()
fyBB.Show
fyBB.ZOrder 0
Me.Enabled = False
End Sub

Private Sub cmdLeft_Click()
Dim tt As String
On Error Resume Next
    mtA.Value = DateSerial(Year(mtA.Value), Month(mtA.Value) - 1, Day(mtA.Value))
If lblFw.Caption = "" Then

        tt = "FydVG('" & mod1.Qy & "','" & frmBxV.mtA.Value & "')"

Else
        If lblFw.ToolTipText = "" Then
            'tt = "FydVGBm2('" & lblFw.Caption & "','" & mod1.Qy & "','" & frmBxV.mtA.Value & "')"
                   tt = "select * from FydBrowG where 部门='" & lblFw.Caption & "' and year(签收日期)=" & Year(frmBxV.mtA.Value) & _
                   " and  (( (month( 签收日期)=" & Month(mtA.Value) & "-1 ) and day(签收日期)>25) or (month(签收日期)=" & Month(mtA) & " and day(签收日期)<26))  order by 签收日期 desc"
                           frmBxV.adoBxV.Close
        frmBxV.adoBxV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Else
            tt = " FydVGywy('" & lblFw.Caption & "','" & frmBxV.mtA.Value & "')"
                    frmBxV.adoBxV.Close
        frmBxV.adoBxV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        End If
End If

        Set frmBxV.dtgBx.DataSource = frmBxV.adoBxV
        If frmBxV.adoBxV.RecordCount > 0 Then
            frmBxV.dtgBx.FixedRows = 0
            frmBxV.dtgBx.MergeCol(1) = True
            frmBxV.dtgBx.MergeCol(2) = True
            frmBxV.dtgBx.MergeCol(3) = True
            frmBxV.dtgBx.MergeCol(4) = True
            frmBxV.dtgBx.MergeCol(5) = True
            frmBxV.dtgBx.MergeCol(7) = True
            frmBxV.dtgBx.MergeCells = 3
            frmBxV.dtgBx.FixedRows = 1
        End If
End Sub

Private Sub cmdNB_Click()
Dim tt As String
On Error Resume Next
tt = "select * from FydBrowG where nlb=81  order by 起始期 desc"
frmBxV.adoBxV.Close
frmBxV.adoBxV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmBxV.dtgBx.DataSource = frmBxV.adoBxV
If frmBxV.adoBxV.RecordCount > 0 Then
    frmBxV.dtgBx.FixedRows = 0
    frmBxV.dtgBx.MergeCol(1) = True
    frmBxV.dtgBx.MergeCol(2) = True
    frmBxV.dtgBx.MergeCol(3) = True
    frmBxV.dtgBx.MergeCol(4) = True
    frmBxV.dtgBx.MergeCol(5) = True
    frmBxV.dtgBx.MergeCol(7) = True
    frmBxV.dtgBx.MergeCells = 3
    frmBxV.dtgBx.FixedRows = 1
End If
End Sub

Private Sub cmdOpen_Click()
dtgBx.Col = 8
'MsgBox MGa.Text

If Val(dtgBx.Text) = 0 Then Exit Sub
If mod1.DKZ(dtgBx.Text, 2) = True Then
        MsgBox "这份表单正由" & mod1.DKRen & "打开,请稍候再试,或与马晓聪联系."
        Exit Sub
End If

frmBxBrow.Enabled = False
frmFYBX.Show

Call ModBx.FyQing
Call ModBx.fydBound(Val(dtgBx.Text))
End Sub

Private Sub cmdRight_Click()
Dim tt As String
On Error Resume Next
mtA.Value = DateSerial(Year(mtA.Value), Month(mtA.Value) + 1, Day(mtA.Value))
If lblFw.Caption = "" Then

        tt = "FydVG('" & mod1.Qy & "','" & frmBxV.mtA.Value & "')"

Else
        If lblFw.ToolTipText = "" Then
            'tt = "FydVGBm2('" & lblFw.Caption & "','" & mod1.Qy & "','" & frmBxV.mtA.Value & "')"
                   tt = "select * from FydBrowG where 部门='" & lblFw.Caption & "' and year(签收日期)=" & Year(frmBxV.mtA.Value) & _
                   " and  (( (month( 签收日期)=" & Month(mtA.Value) & "-1 ) and day(签收日期)>25) or (month(签收日期)=" & Month(mtA) & " and day(签收日期)<26))  order by 签收日期 desc"
                           frmBxV.adoBxV.Close
        frmBxV.adoBxV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
        Else
            tt = " FydVGywy('" & lblFw.Caption & "','" & frmBxV.mtA.Value & "')"
            frmBxV.adoBxV.Close
            frmBxV.adoBxV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        End If
End If

        Set frmBxV.dtgBx.DataSource = frmBxV.adoBxV
        If frmBxV.adoBxV.RecordCount > 0 Then
            frmBxV.dtgBx.FixedRows = 0
            frmBxV.dtgBx.MergeCol(1) = True
            frmBxV.dtgBx.MergeCol(2) = True
            frmBxV.dtgBx.MergeCol(3) = True
            frmBxV.dtgBx.MergeCol(4) = True
            frmBxV.dtgBx.MergeCol(5) = True
            frmBxV.dtgBx.MergeCol(7) = True
            frmBxV.dtgBx.MergeCells = 3
            frmBxV.dtgBx.FixedRows = 1
        End If
End Sub

Private Sub cmdV_Click()
Dim tt As String
Dim qq As String
On Error Resume Next
If comLx.Text = "金额" Then
    tt = "select * from FydBrowG where 金额=" & Val(txtZ.Text)
    
ElseIf comLx.Text = "编号" Then
    tt = "select * from FydBrowG where 编号=" & Val(txtZ.Text)
ElseIf comLx.Text = "报销人" Then
    tt = "select * from FydBrowG where 报销人 like '%" & Trim(txtZ.Text) & "%'"
ElseIf comLx.Text = "报销内容" Then
    tt = "select * from FydBrowGG where bz like '%" & Trim(txtZ.Text) & "%' or khmc like '%" & Trim(txtZ.Text) & "%'"
End If
If mod1.Qy <> "上海" Then
    qq = " and 区域='" & mod1.Qy & "' order by 签收日期 desc,报销人"
Else
    qq = " order by 签收日期 desc,报销人"
End If
tt = tt & qq
frmBxV.adoBxV.Close
frmBxV.adoBxV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set frmBxV.dtgBx.DataSource = frmBxV.adoBxV
If frmBxV.adoBxV.RecordCount > 0 Then
    frmBxV.dtgBx.FixedRows = 0
    frmBxV.dtgBx.MergeCol(1) = True
    frmBxV.dtgBx.MergeCol(2) = True
    frmBxV.dtgBx.MergeCol(3) = True
    frmBxV.dtgBx.MergeCol(4) = True
    frmBxV.dtgBx.MergeCol(5) = True
    frmBxV.dtgBx.MergeCol(7) = True
    frmBxV.dtgBx.MergeCells = 3
    frmBxV.dtgBx.FixedRows = 1
End If
End Sub

Private Sub cmdWb_Click()
Dim tt As String
On Error Resume Next
If mod1.Bm = "商务部" And mod1.Qy = "上海" Then
    tt = "fydVw"
ElseIf mod1.Bq2 = True And mod1.Qy <> "上海" Then '外地办文员
    tt = "fydVwQy('" & mod1.Qy & "')"
ElseIf mod1.BmJl = True And mod1.Bm <> "商务部" Then

ElseIf mod1.DName = "宋晓炯" Or mod1.DName = "宋晓炯1" Then  '总经理
    tt = "fydVwcomid(" & mod1.comId & ")"
End If
frmBxV.adoBxV.Close
frmBxV.adoBxV.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
Set frmBxV.dtgBx.DataSource = frmBxV.adoBxV
If frmBxV.adoBxV.RecordCount > 0 Then
    frmBxV.dtgBx.FixedRows = 0
    frmBxV.dtgBx.MergeCol(1) = True
    frmBxV.dtgBx.MergeCol(2) = True
    frmBxV.dtgBx.MergeCol(3) = True
    frmBxV.dtgBx.MergeCol(4) = True
    frmBxV.dtgBx.MergeCol(5) = True
    frmBxV.dtgBx.MergeCol(7) = True
    frmBxV.dtgBx.MergeCells = 3
    frmBxV.dtgBx.FixedRows = 1
End If
End Sub

Private Sub dtgBx_DblClick()
Static Px As Boolean

If dtgBx.Row = 1 Then
    If Px = True Then
        dtgBx.Sort = 2
        Px = False
    Else
        dtgBx.Sort = 1
        Px = True
    End If

End If
End Sub


Private Sub dtgBx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static ZF As Boolean
If Button <> 2 Then Exit Sub
If ZF = False Then
        dtgBx.FixedRows = 0
        dtgBx.MergeCells = 0
        dtgBx.FixedRows = 1
        ZF = True
Else
        dtgBx.FixedRows = 0
        dtgBx.MergeCells = 3
        dtgBx.FixedRows = 1
        ZF = False
End If
End Sub


Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Set adoBxV = New ADODB.Recordset
dtgBx.ColWidth(0) = 300
dtgBx.ColWidth(9) = 0
dtgBx.ColWidth(10) = 0
dtgBx.ColWidth(11) = 0
If mod1.KhK = 2 Or mod1.KhK = 3 Or mod1.Bm = "商务部" Then
    cmdNB.Visible = True
Else
    cmdNB.Visible = False
End If
If mod1.KhK = 1 Then
    cmdCWBB.Visible = False
End If


dtgDZ.Cols = 5
dtgDZ.Rows = 100
dtgDZ.Row = 0
dtgDZ.Col = 0: dtgDZ.Text = "公告日期": dtgDZ.CellFontBold = True
dtgDZ.Col = 1: dtgDZ.Text = "客户名称": dtgDZ.CellFontBold = True
dtgDZ.Col = 2: dtgDZ.Text = "金额": dtgDZ.CellFontBold = True
dtgDZ.Col = 3: dtgDZ.Text = "到帐日期": dtgDZ.CellFontBold = True
dtgDZ.Col = 4: dtgDZ.Text = "备注": dtgDZ.CellFontBold = True
dtgDZ.ColWidth(1) = 2500
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
Cancel = True
frmZu.TBa.Buttons(3).Value = tbrUnpressed
'frmBxBrow.WindowState = 0
frmZu.Enabled = True
frmBxV.Visible = False
End If
End Sub


Private Sub lblFw_DblClick()
lblFw.Caption = ""
lblFw.ToolTipText = ""
End Sub

