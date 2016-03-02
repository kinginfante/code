VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGzbN 
   BackColor       =   &H00C0FFC0&
   Caption         =   "销售工作报告"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "费用明细"
      Height          =   225
      Left            =   7320
      TabIndex        =   23
      Top             =   8850
      Width           =   1395
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "合同评审"
      Height          =   255
      Left            =   5520
      TabIndex        =   22
      Top             =   8850
      Width           =   1485
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "工作报告"
      Height          =   255
      Left            =   3840
      TabIndex        =   21
      Top             =   8850
      Width           =   1395
   End
   Begin VB.CommandButton cmdXZ 
      Caption         =   "选择"
      Height          =   555
      Left            =   12510
      Picture         =   "frmGzbN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "选择人员"
      Top             =   8610
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   555
      Left            =   8130
      TabIndex        =   18
      Top             =   6090
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   979
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制"
      Height          =   555
      Left            =   13890
      Picture         =   "frmGzbN.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "点击后,打开EXCEL,可将表格内容粘贴."
      Top             =   8610
      Width           =   675
   End
   Begin VB.CommandButton cmdKZ 
      Caption         =   "视角"
      Height          =   555
      Left            =   13200
      Picture         =   "frmGzbN.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8610
      Width           =   675
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      TabIndex        =   6
      Top             =   8100
      Width           =   15255
      Begin VB.Label lblJZ 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   13920
         TabIndex        =   15
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "价值(收款-费用)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   11430
         TabIndex        =   14
         Top             =   60
         Width           =   2055
      End
      Begin VB.Label lblSK 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   8910
         TabIndex        =   13
         Top             =   60
         Width           =   1035
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "收款(完成指标)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   225
         Left            =   6750
         TabIndex        =   12
         Top             =   60
         Width           =   1725
      End
      Begin VB.Label lblYS 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5010
         TabIndex        =   11
         Top             =   60
         Width           =   1125
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "应收"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   60
         Width           =   555
      End
      Begin VB.Label lblFy 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   2250
         TabIndex        =   9
         Top             =   60
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "费用"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   225
         Left            =   1470
         TabIndex        =   8
         Top             =   60
         Width           =   585
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   7
         Top             =   60
         Width           =   675
      End
   End
   Begin MSComCtl2.DTPicker dtPDate 
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   8850
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   450
      _Version        =   393216
      CalendarForeColor=   32768
      CustomFormat    =   "yyyy年"
      Format          =   106102787
      CurrentDate     =   39968
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgB 
      Height          =   8085
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   14261
      _Version        =   393216
      BackColor       =   12648384
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      WordWrap        =   -1  'True
      ScrollTrack     =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00C0FFFF&
      Caption         =   "返回"
      Height          =   555
      Left            =   14580
      Picture         =   "frmGzbN.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8610
      Width           =   645
   End
   Begin VB.Label lblRen 
      BackStyle       =   0  'Transparent
      Caption         =   "lblRen"
      Height          =   255
      Left            =   11010
      TabIndex        =   19
      Top             =   8820
      Width           =   1245
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   15240
      Y1              =   8550
      Y2              =   8550
   End
   Begin VB.Label lblLR 
      Caption         =   "lblLR"
      Height          =   225
      Left            =   2160
      TabIndex        =   5
      Top             =   8880
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Label lblFr 
      Caption         =   "lblFr"
      Height          =   165
      Left            =   1920
      TabIndex        =   4
      Top             =   8640
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "年度"
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   8610
      Width           =   1155
   End
End
Attribute VB_Name = "frmGzbN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FS As Boolean '表格视角
Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
End Sub



Private Sub cmdCopy_Click()
Dim oo As Integer: Dim ii As Integer

dtgN.Rows = dtgB.Rows: dtgN.Cols = dtgB.Cols
For oo = 0 To dtgB.Rows - 1
    dtgB.Row = oo: dtgN.Row = oo
    For ii = 0 To dtgB.Cols - 1
        dtgB.Col = ii: dtgN.Col = ii
        If ii = 3 Or ii = 5 Then
            dtgN.Text = Replace(dtgB.Text, Chr(13), "")
            dtgN.Text = Replace(dtgN.Text, Chr(10), "")

        Else
            dtgN.Text = dtgB.Text
        End If
    Next
Next
dtgN.Col = 0
dtgN.Row = 0
dtgN.ColSel = 6
dtgN.RowSel = dtgN.Rows - 1
Clipboard.Clear
Clipboard.SetText dtgN.Clip
End Sub

Private Sub cmdKZ_Click()
Call QV(FS)
End Sub

Private Sub cmdXZ_Click()
Set Ren.XForm = New frmGzbN
Call mod1.RenXz("frmGzbN", Me, 0)
End Sub

Private Sub dtgB_DblClick()
Dim Frq As Date
Call frmGZbN1.Qing
Call frmGZbN1.Qing
dtgB.Col = 0
Frq = Left(dtgB.Text, 10)
Call frmGZbN1.Bound(lblRen.ToolTipText, Frq)
frmGZbN1.Show
frmGZbN1.txtBz4.SetFocus
End Sub

Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0: Me.Top = 0
dtgB.Cols = 8

dtgB.ColWidth(0) = 2535

dtgB.ColWidth(1) = 1005
dtgB.ColWidth(3) = 4050
dtgB.ColWidth(5) = 4230
dtgB.ColWidth(7) = 0
lblRen.Caption = ""
End Sub

Public Sub WeekDate(Ddate As Date, Uid As String)
Dim tt As String
Dim oo As Long: Dim ii As Long: Dim bb As Long: Dim cc As Long: Dim DD As Long: Dim EE As Long
Dim TFrq As Date: Dim TLrq As Date
Dim Ra, La, Rb, Lb, RC, Lc, RD, Ld
Dim DH As Single
frmWait.Show
frmWait.ZOrder 0
frmWait.Refresh
dtgB.Visible = False
dtgB.Clear
For oo = 1 To dtgB.Rows - 1
    dtgB.RowHeight(oo) = dtgB.RowHeight(0)
Next
dtgB.Rows = 56
dtgB.Row = 0
dtgB.Col = 0: dtgB.Text = "日期": dtgB.CellFontBold = True
dtgB.Col = 1: dtgB.Text = "费用": dtgB.CellFontBold = True
dtgB.Col = 2: dtgB.Text = "速达费用": dtgB.CellFontBold = True: dtgB.CellForeColor = &H8000&
dtgB.Col = 3: dtgB.Text = "应收客户": dtgB.CellFontBold = True
dtgB.Col = 4: dtgB.Text = "应收款": dtgB.CellFontBold = True
dtgB.Col = 5: dtgB.Text = "收款客户": dtgB.CellFontBold = True: dtgB.CellForeColor = &H8000&
dtgB.Col = 6: dtgB.Text = "实际收款": dtgB.CellFontBold = True: dtgB.CellForeColor = &H8000&

Call GetWeek(Ddate)
dtgB.Row = 1: dtgB.Col = 0

For oo = 1 To dtgB.Rows
  dtgB.Row = oo
  dtgB.Text = Format(lblFr.Caption, "YYYY-MM-dd") & "  To  " & Format(lblLr.Caption, "YYYY-MM-dd")

  lblLr.Caption = DateSerial(Year(lblFr.Caption), Month(lblFr.Caption), Day(lblFr.Caption) - 1)
  lblFr.Caption = DateSerial(Year(lblLr.Caption), Month(lblLr.Caption), Day(lblLr.Caption) - 6)
  If lblLr.Caption < DateSerial(Year(Ddate), 4, 1) Then
    Exit For
  End If
Next
On Error GoTo frmgzbnERR2
tt = "SELECT dbo.htping1.rq, dbo.htping1.yingfJe, dbo.htPing.htBh, dbo.htPing.xmMc " & _
        "FROM dbo.htping1 INNER JOIN dbo.htPing ON dbo.htping1.htBh = dbo.htPing.Hid " & _
        "WHERE (YEAR(dbo.htping1.rq) > 2005) AND (dbo.htPing.DelF = 1) AND (dbo.htPing.htF = 1 OR dbo.htPing.htF = 2 OR dbo.htPing.htF = 9) and dbo.htping.xuid='" & Uid & _
        "' and dbo.htping1.rq>='" & DateSerial(Year(dtpDate.Value), 4, 1) & "' and dbo.htping1.rq<'" & DateSerial(Year(dtpDate.Value) + 1, 4, 1) & "' order by dbo.htping1.rq desc;" & _
        "SELECT  SD30301_豪曼制冷.dbo.b_expense.billdate,SD30301_豪曼制冷.dbo.b_expensedetail.amount" & _
        " FROM SD30301_豪曼制冷.dbo.b_expense INNER JOIN SD30301_豪曼制冷.dbo.b_expensedetail ON" & _
        " SD30301_豪曼制冷.dbo.b_expense.billid = SD30301_豪曼制冷.dbo.b_expensedetail.billid INNER JOIN" & _
        " SD30301_豪曼制冷.dbo.l_employ ON SD30301_豪曼制冷.dbo.b_expense.empid = SD30301_豪曼制冷.dbo.l_employ.empid INNER JOIN" & _
        " SD30301_豪曼制冷.dbo.l_department ON SD30301_豪曼制冷.dbo.b_expense.departmentid = SD30301_豪曼制冷.dbo.l_department.departmentid" & _
        " where SD30301_豪曼制冷.dbo.l_employ.code='" & Right(Uid, 3) & "' and  SD30301_豪曼制冷.dbo.b_expense.billdate>='" & DateSerial(Year(Ddate), 4, 1) & _
        "' and SD30301_豪曼制冷.dbo.b_expense.billdate<'" & DateSerial(Year(Ddate) + 1, 4, 1) & "' order by SD30301_豪曼制冷.dbo.b_expense.billdate desc;" & _
        " select khmc,htbh,billdate,amount from SDV_ChargeA where code='" & Uid & "' and billdate>='" & DateSerial(Year(Ddate), 4, 1) & _
        "' and billdate<'" & DateSerial(Year(Ddate) + 1, 4, 1) & "' order by billdate desc;" & _
        " select bf,fr from SalesReport where uid='" & Uid & "' order by fr desc"
        
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
If mod1.HTP.BOF = False Then
Ra = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
Rb = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
RC = mod1.HTP.GetRows
End If
Set mod1.HTP = mod1.HTP.NextRecordset
If mod1.HTP.BOF = False Then
RD = mod1.HTP.GetRows
End If
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next
La = UBound(Ra, 2) + 1
Lb = UBound(Rb, 2) + 1
Lc = UBound(RC, 2) + 1
Ld = UBound(RD, 2) + 1
bb = 0: cc = 0: DD = 0
lblFy.Caption = 0: lblYs.Caption = 0: lblSK.Caption = 0: lblJZ.Caption = 0
For oo = 1 To dtgB.Rows - 1
    dtgB.Row = oo
    dtgB.Col = 0
    If dtgB.Text = "" Then Exit For
    TFrq = Left(Trim(dtgB.Text), 10)
    TLrq = Right(Trim(dtgB.Text), 10)
    If La > 0 Then
        For ii = bb To La
            If Ra(0, ii) >= TFrq And Ra(0, ii) <= TLrq Then
                dtgB.Col = 3
                dtgB.Text = dtgB.Text & Ra(3, ii) & ":" & Chr(13) & Chr(10) & Ra(2, ii) & Chr(13) & Chr(10)

'''''                DH = 255 * mod1.HH(dtgB.Text, UpInt(dtgB.CellWidth / 200)) / 2
'''''                If DH > dtgB.RowHeight(dtgB.Row) Then
'''''                    dtgB.RowHeight(dtgB.Row) = DH
'''''                End If
                dtgB.Col = 4
                dtgB.Text = Val(dtgB.Text) + Ra(1, ii)
                lblYs.Caption = Round(Val(lblYs.Caption) + Ra(1, ii), 0)
                bb = ii
            ElseIf Ra(0, ii) < TFrq Then
                bb = ii
                Exit For
            End If
        Next
    End If
    If Lb > 0 Then
        For ii = cc To Lb
            If Rb(0, ii) >= TFrq And Rb(0, ii) <= TLrq Then
                dtgB.Col = 2
                dtgB.Text = Val(dtgB.Text) + Rb(1, ii)
                lblFy.Caption = Round(Val(lblFy.Caption) + Rb(1, ii), 0)
                cc = ii
                dtgB.CellForeColor = &H8000&
            ElseIf Rb(0, ii) < TFrq Then
                cc = ii
                Exit For
            End If
        Next
    End If
    If Lc > 0 Then
        For ii = DD To Lc
            If RC(2, ii) >= TFrq And RC(2, ii) <= TLrq Then
                dtgB.Col = 5
                dtgB.Text = dtgB.Text & RC(0, ii) & ":" & Chr(13) & Chr(10) & RC(1, ii) & ":" & Chr(13) & Chr(10)
'''''                DH = 255 * mod1.HH(dtgB.Text, UpInt(dtgB.CellWidth / 200)) / 2
'''''                If DH > dtgB.RowHeight(dtgB.Row) Then
'''''                    dtgB.RowHeight(dtgB.Row) = DH
'''''                End If
                dtgB.CellForeColor = &H8000&
                dtgB.Col = 6
                dtgB.Text = Round(Val(dtgB.Text) + RC(3, ii), 0)
                lblSK.Caption = Round(Val(lblSK.Caption) + RC(3, ii), 0)
                DD = ii
                dtgB.CellForeColor = &H8000&
            ElseIf RC(2, ii) < TFrq Then
                DD = ii
                Exit For
            End If
        Next
    End If
    If Ld > 0 Then
        For ii = EE To Ld
            If RD(1, ii) = TFrq Then
                dtgB.Col = 1
                dtgB.Text = RD(0, ii)
                EE = ii
            ElseIf RD(1, ii) < TFrq Then
                EE = ii
                Exit For
            End If
        Next
    End If


Next
lblJZ.Caption = Val(lblSK.Caption) - Val(lblFy.Caption)
dtgB.Visible = True
frmWait.Visible = False
'''''''If La > 0 Then
'''''''    bb = 0
'''''''    For oo = 1 To dtgB.Rows - 1
'''''''        dtgB.Row = oo
'''''''        dtgB.Col = 0
'''''''        If dtgB.Text = "" Then Exit For
'''''''        TFrq = Left(Trim(dtgB.Text), 10)
'''''''        TLrq = Right(Trim(dtgB.Text), 10)
'''''''        For ii = bb To La
'''''''            If Ra(0, ii) >= TFrq And Ra(0, ii) <= TLrq Then
'''''''                dtgB.Col = 3
'''''''                dtgB.Text = dtgB.Text & Ra(3, ii) & ":" & Ra(2, ii) & ":"
'''''''                dtgB.Col = 4
'''''''                dtgB.Text = Val(dtgB.Text) + Ra(1, ii)
'''''''                bb = ii
'''''''            ElseIf Ra(0, ii) < TFrq Then
'''''''                bb = ii
'''''''                Exit For
'''''''            End If
'''''''        Next
'''''''        If bb = La Then
'''''''            Exit For
'''''''        End If
'''''''
'''''''    Next
'''''''End If
'''''''If Lb > 0 Then
'''''''    bb = 0
'''''''    For oo = 1 To dtgB.Rows - 1
'''''''        dtgB.Row = oo
'''''''        dtgB.Col = 0
'''''''        If dtgB.Text = "" Then Exit For
'''''''        TFrq = Left(Trim(dtgB.Text), 10)
'''''''        TLrq = Right(Trim(dtgB.Text), 10)
'''''''        For ii = bb To Lb
'''''''            If Rb(0, ii) >= TFrq And Rb(0, ii) <= TLrq Then
'''''''                dtgB.Col = 2
'''''''                dtgB.Text = Val(dtgB.Text) + Rb(1, ii)
'''''''                bb = ii
'''''''            ElseIf Rb(0, ii) < TFrq Then
'''''''                bb = ii
'''''''                Exit For
'''''''            End If
'''''''        Next
'''''''        If bb = Lb Then
'''''''            Exit For
'''''''        End If
'''''''
'''''''    Next
'''''''End If
'''''''If Lc > 0 Then
'''''''    bb = 0
'''''''    For oo = 1 To dtgB.Rows - 1
'''''''        dtgB.Row = oo
'''''''        dtgB.Col = 0
'''''''        If dtgB.Text = "" Then Exit For
'''''''        TFrq = Left(Trim(dtgB.Text), 10)
'''''''        TLrq = Right(Trim(dtgB.Text), 10)
'''''''        For ii = bb To Lc
'''''''            If Rc(2, ii) >= TFrq And Rc(2, ii) <= TLrq Then
'''''''                dtgB.Col = 5
'''''''                dtgB.Text = dtgB.Text & Rc(0, ii) & ":" & Rc(1, ii) & ":"
'''''''                dtgB.Col = 6
'''''''                dtgB.Text = Val(dtgB.Text) + Rc(3, ii)
'''''''                bb = ii
'''''''            ElseIf Rc(2, ii) < TFrq Then
'''''''                bb = ii
'''''''                Exit For
'''''''            End If
'''''''        Next
'''''''        If bb = Lc Then
'''''''            Exit For
'''''''        End If
'''''''
'''''''    Next
'''''''End If
Exit Sub
frmgzbnERR2:
MsgBox "出错!"
End Sub
Public Sub GetWeek(mtA As Date)
Select Case DatePart("w", mtA)
Case 1 '星期日
lblFr.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) - 6)
lblLr.Caption = mtA
Case 2 '星期一
lblFr.Caption = mtA
lblLr.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) + 6)
Case 3
lblFr.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) - 1)
lblLr.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) + 5)
Case 4
lblFr.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) - 2)
lblLr.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) + 4)
Case 5
lblFr.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) - 3)
lblLr.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) + 3)
Case 6
lblFr.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) - 4)
lblLr.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) + 2)
Case 7
lblFr.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) - 5)
lblLr.Caption = DateSerial(Year(mtA), Month(mtA), Day(mtA) + 1)
End Select
End Sub

Private Sub Label6_Click()

End Sub



Public Sub QV(FS As Boolean)
Dim oo As Integer
Dim DH As Single
dtgB.Visible = False
For oo = 1 To dtgB.Rows - 1
    dtgB.Row = oo
    dtgB.Col = 0
    If dtgB.Text = "" Then
        Exit For
    End If
    If FS = True Then
        dtgB.Col = 3
        DH = 255 * mod1.HH(dtgB.Text, UpInt(dtgB.CellWidth / 200)) / 2
        If DH > dtgB.RowHeight(dtgB.Row) Then
            dtgB.RowHeight(dtgB.Row) = DH
        End If
        dtgB.Col = 5
        DH = 255 * mod1.HH(dtgB.Text, UpInt(dtgB.CellWidth / 200)) / 2
        If DH > dtgB.RowHeight(dtgB.Row) Then
            dtgB.RowHeight(dtgB.Row) = DH
        End If
    Else
        dtgB.RowHeight(oo) = dtgB.RowHeight(0)
    End If

Next
    If FS = True Then
        Me.FS = False
    Else
        Me.FS = True
    End If
dtgB.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = True
Cancel = True
frmZu.Enabled = True
End Sub


