VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmPeiView 
   BackColor       =   &H00C0FFC0&
   Caption         =   "培训汇总表"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查询所有"
      Height          =   315
      Left            =   4380
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8760
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgDN 
      Height          =   465
      Left            =   60
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   820
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   555
      Left            =   7170
      TabIndex        =   15
      Top             =   8460
      Width           =   4815
      Begin VB.OptionButton opt3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "年度"
         Height          =   375
         Left            =   3150
         TabIndex        =   18
         Top             =   210
         Width           =   1215
      End
      Begin VB.OptionButton opt2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "季度"
         Height          =   375
         Left            =   1695
         TabIndex        =   17
         Top             =   210
         Width           =   1215
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "月度"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   210
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14490
      Picture         =   "PeiView.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8490
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   315
      Left            =   4740
      TabIndex        =   13
      Top             =   8820
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">"
      Height          =   285
      Left            =   1710
      TabIndex        =   12
      Top             =   8790
      Width           =   375
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<"
      Height          =   285
      Left            =   630
      TabIndex        =   10
      Top             =   8790
      Width           =   375
   End
   Begin VB.CommandButton cmdZZ 
      BackColor       =   &H00C0C0FF&
      Caption         =   "人员部门选择"
      Height          =   315
      Left            =   2670
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8760
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8385
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   14790
      _Version        =   393216
      BackColor       =   12648384
      Rows            =   30
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgDe 
      Height          =   7845
      Left            =   6120
      TabIndex        =   1
      Top             =   540
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   13838
      _Version        =   393216
      BackColor       =   16777152
      Rows            =   30
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   12648447
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.Label lblYear 
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      Height          =   285
      Left            =   1110
      TabIndex        =   11
      Top             =   8820
      Width           =   585
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "年限"
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   8820
      Width           =   435
   End
   Begin VB.Label lblZw 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   11730
      TabIndex        =   7
      Top             =   150
      Width           =   1485
   End
   Begin VB.Label lblBh 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   9390
      TabIndex        =   6
      Top             =   150
      Width           =   1485
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   6900
      TabIndex        =   5
      Top             =   150
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "职务："
      Height          =   285
      Left            =   11040
      TabIndex        =   4
      Top             =   150
      Width           =   675
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "工号："
      Height          =   285
      Left            =   8670
      TabIndex        =   3
      Top             =   150
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "姓名："
      Height          =   285
      Left            =   6240
      TabIndex        =   2
      Top             =   150
      Width           =   645
   End
End
Attribute VB_Name = "frmPeiView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub dtgbrFF()
dtgBr.Clear
dtgBr.Cols = 7: dtgBr.Rows = 50
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "员工姓名": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "参训次数": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "参训总时数": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "参训出勤率": dtgBr.CellFontBold = True
dtgBr.Col = 4: dtgBr.Text = "培训总费用": dtgBr.CellFontBold = True
dtgBr.Col = 5: dtgBr.Text = "UserZw": dtgBr.CellFontBold = True
dtgBr.Col = 6: dtgBr.Text = "ID": dtgBr.CellFontBold = True
dtgBr.ColWidth(5) = 0
dtgBr.ColWidth(6) = 0
dtgBr.ColWidth(2) = 1200
dtgBr.ColWidth(3) = 1200
dtgBr.ColWidth(4) = 1200

dtgN.Clear
dtgN.Cols = 7: dtgN.Rows = 50
dtgN.Row = 0
dtgN.Col = 0: dtgN.Text = "员工姓名": dtgN.CellFontBold = True
dtgN.Col = 1: dtgN.Text = "参训次数": dtgN.CellFontBold = True
dtgN.Col = 2: dtgN.Text = "参训总时数": dtgN.CellFontBold = True
dtgN.Col = 3: dtgN.Text = "参训出勤率": dtgN.CellFontBold = True
dtgN.Col = 4: dtgN.Text = "培训总费用": dtgN.CellFontBold = True
dtgN.Col = 5: dtgN.Text = "UserZw": dtgN.CellFontBold = True
dtgN.Col = 6: dtgN.Text = "ID": dtgN.CellFontBold = True
dtgN.ColWidth(5) = 0
dtgN.ColWidth(6) = 0
dtgN.ColWidth(2) = 1200
dtgN.ColWidth(3) = 1200
dtgN.ColWidth(4) = 1200
End Sub

Private Sub cmdAll_Click()
        tt = "select name,cq,cq1,zt,zfy,userZw,uid from peiView30 order by bm"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        dtgBr.Visible = False
        Call frmPeiView.BRBound(Ra)
        dtgBr.Visible = True
        
        Frame1.Visible = False
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub cmdLeft_Click()
Dim tt As String
Dim Ra
Dim Rb
lblYear.Caption = Val(lblYear.Caption) - 1

tt = "select name,cq,cq1,zt,zfy,userZw,uid from peiView3 where bm='" & Ren.lblBM.Caption & "' and nd=" & Val(frmPeiView.lblYear.Caption) & _
    "select username from worker where zzf=1 and bm='" & Ren.lblBM.Caption & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Call frmPeiView.BRBMBound(Ra, Rb)
Call dtgDeFF
End Sub

Private Sub cmdRight_Click()
Dim tt As String
Dim Ra
Dim Rb
lblYear.Caption = Val(lblYear.Caption) + 1

tt = "select name,cq,cq1,zt,zfy,userZw,uid from peiView3 where bm='" & Ren.lblBM.Caption & "' and nd=" & Val(frmPeiView.lblYear.Caption) & _
    "select username from worker where zzf=1 and bm='" & Ren.lblBM.Caption & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Call frmPeiView.BRBMBound(Ra, Rb)
Call dtgDeFF
End Sub

Private Sub cmdZZ_Click()
Frame1.Visible = True
    Call mod1.RenXz("frmPeiView", Me, 0)

End Sub

Private Sub dtgBr_Click()
Dim tt As String
Dim ii As Integer
Dim jj As Integer '月季年
Dim NN As Integer
Dim Fy As Single
Dim PxT As Single
Dim YY As Integer
Dim Ra
Dim La As Long
Dim oo As Long
Dim NI As Long '序号
dtgN.Row = dtgBr.Row
If dtgN.Row = 0 Then Exit Sub
dtgN.Col = 0: lblName.Caption = dtgN.Text
dtgN.Col = 6: lblBh.Caption = dtgN.Text
dtgN.Col = 5: lblZw.Caption = dtgN.Text
If Frame1.Visible = True Then
    tt = "select zid,lb,mc,ft,adr,cf,pxt,fy,bz from peiView5 where uid='" & lblBh.Caption & "' and nd=" & Val(lblYear.Caption) & " order by ft"
Else '查询所有
    tt = "select zid,lb,mc,ft,adr,cf,pxt,fy,bz from peiView5 where uid='" & lblBh.Caption & "'  order by ft desc"
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgDe.Visible = False
dtgDe.Rows = La + 50
Call dtgDeFF
For oo = 1 To La
    dtgDe.Row = oo
    dtgDe.Col = 0: dtgDe.Text = Ra(0, oo - 1): dtgDe.CellForeColor = &H80000012
    dtgDe.Col = 1: dtgDe.Text = Ra(1, oo - 1): dtgDe.CellForeColor = &H80000012
    dtgDe.Col = 2: dtgDe.Text = Ra(2, oo - 1): dtgDe.CellForeColor = &H80000012
    dtgDe.Col = 3: dtgDe.Text = Ra(3, oo - 1): dtgDe.CellForeColor = &H80000012
    dtgDe.Col = 4: dtgDe.Text = Ra(4, oo - 1): dtgDe.CellForeColor = &H80000012
    dtgDe.Col = 5: dtgDe.Text = Ra(5, oo - 1): dtgDe.CellForeColor = &H80000012
    dtgDe.Col = 6: dtgDe.Text = Ra(6, oo - 1): dtgDe.CellForeColor = &H80000012
    dtgDe.Col = 7: dtgDe.Text = Round(Ra(7, oo - 1), 2): dtgDe.CellForeColor = &H80000012
    dtgDe.Col = 8: dtgDe.Text = Ra(8, oo - 1): dtgDe.CellForeColor = &H80000012
    dtgDe.RowHeight(oo) = dtgDe.RowHeight(0) * 3

Next
If Frame1.Visible = False Then '查询所有
    For oo = 1 To La
        NI = oo
        dtgDe.Row = oo
        dtgDe.Col = 5
        If dtgDe.Text = "True" Then
            dtgDe.Text = "出勤"
        ElseIf dtgDe.Text = "False" Then
            dtgDe.Text = "缺勤"
            For ii = 0 To 8
                dtgDe.Col = ii
                dtgDe.CellForeColor = &HFF&
            Next
        End If
        dtgDe.Col = 0: dtgDe.Text = NI
    Next
        dtgDe.Visible = True
    Exit Sub
End If
'表格统计
dtgDN.Clear: dtgDN.Rows = dtgDe.Rows: dtgDN.Cols = dtgDe.Cols
If opt1.Value = True Then
    jj = 1: NN = 1: Fy = 0: PxT = 0
    For oo = 1 To La + 30
        dtgDe.Row = oo
        dtgDe.Col = 0
        If dtgDe.Text = "" Then Exit For
        dtgDe.Col = 3
        If Month(dtgDe.Text) = jj Then '累加
             dtgDe.Col = 7: Fy = Fy + Val(dtgDe.Text)
             dtgDe.Col = 5
             If dtgDe.Text = "True" Then
                dtgDe.Col = 6: PxT = PxT + Val(dtgDe.Text)
             End If
             dtgDN.Row = NN
             For ii = 0 To 8
                dtgDe.Col = ii: dtgDN.Col = ii
                dtgDN.Text = dtgDe.Text
             Next
             NN = NN + 1
        Else
            dtgDN.Row = NN
            dtgDN.Col = 2: dtgDN.Text = Str(jj) & "月小计:"
            dtgDN.Col = 6: dtgDN.Text = PxT
            dtgDN.Col = 7: dtgDN.Text = Fy
            Fy = 0: PxT = 0
            NN = NN + 1
            jj = jj + 1
            oo = oo - 1
        End If
    Next
    For YY = jj To 12
            dtgDN.Row = NN
            dtgDN.Col = 1: dtgDN.Text = Str(YY) & "月小计:"
            dtgDN.Col = 6: dtgDN.Text = PxT
            dtgDN.Col = 7: dtgDN.Text = Fy
            NN = NN + 1
            PxT = 0: Fy = 0
    Next
    'dtgDe.Clear
    '回拷
    NI = 1
    For oo = 1 To dtgDN.Rows
        dtgDe.Row = oo: dtgDN.Row = oo
        For ii = 0 To 8
            dtgDe.Col = ii: dtgDN.Col = ii
            dtgDe.Text = dtgDN.Text
        Next
        dtgDe.Col = 0
        If dtgDe.Text = "" Then
            dtgDe.RowHeight(oo) = dtgDe.RowHeight(0)
        Else
            dtgDe.RowHeight(oo) = dtgDe.RowHeight(0) * 3
        End If
        dtgDe.Col = 5
        If dtgDe.Text = "True" Then
            dtgDe.Text = "出勤"
        ElseIf dtgDe.Text = "False" Then
            dtgDe.Text = "缺勤"
            For ii = 0 To 8
                dtgDe.Col = ii
                dtgDe.CellForeColor = &HFF&
            Next
        End If
        dtgDe.Col = 0
        If Val(dtgDe.Text) > 0 Then
            dtgDe.Text = NI
            NI = NI + 1
        Else
            NI = 1
        End If
    Next
ElseIf opt2.Value = True Then
    jj = 1: NN = 1: Fy = 0: PxT = 0
    For oo = 1 To La + 30
        dtgDe.Row = oo
        dtgDe.Col = 0
        If dtgDe.Text = "" Then Exit For
        dtgDe.Col = 3
        If Month(dtgDe.Text) / 3 <= jj Then '累加
             dtgDe.Col = 7: Fy = Fy + Val(dtgDe.Text)
             dtgDe.Col = 5
             If dtgDe.Text = "True" Then
                dtgDe.Col = 6: PxT = PxT + Val(dtgDe.Text)
             End If
             dtgDN.Row = NN
             For ii = 0 To 8
                dtgDe.Col = ii: dtgDN.Col = ii
                dtgDN.Text = dtgDe.Text
             Next
             NN = NN + 1
        Else
            dtgDN.Row = NN
            dtgDN.Col = 1: dtgDN.Text = Trim(Str(jj)) & "季度统计:"
            dtgDN.Col = 6: dtgDN.Text = PxT
            dtgDN.Col = 7: dtgDN.Text = Fy
            Fy = 0: PxT = 0
            NN = NN + 1
            jj = jj + 1
            oo = oo - 1
        End If
    Next
    For YY = jj To 4
            dtgDN.Row = NN
            dtgDN.Col = 2: dtgDN.Text = Trim(Str(YY)) & "季度统计:"
'''             dtgDe.Col = 5
'''             If dtgDe.Text = "True" Then
'''                dtgDe.Col = 6: PxT = PxT + Val(dtgDe.Text)
'''             End If
            dtgDN.Col = 6: dtgDN.Text = PxT
            dtgDN.Col = 7: dtgDN.Text = Fy
            NN = NN + 1
            PxT = 0: Fy = 0
    Next
    'dtgDe.Clear
    '回拷
    NI = 1
    For oo = 1 To dtgDN.Rows
        dtgDe.Row = oo: dtgDN.Row = oo
        For ii = 0 To 8
            dtgDe.Col = ii: dtgDN.Col = ii
            dtgDe.Text = dtgDN.Text
        Next
        dtgDe.Col = 0
        If dtgDe.Text = "" Then
            dtgDe.RowHeight(oo) = dtgDe.RowHeight(0)
        Else
            dtgDe.RowHeight(oo) = dtgDe.RowHeight(0) * 3
        End If
        dtgDe.Col = 5
        If dtgDe.Text = "True" Then
            dtgDe.Text = "出勤"
        ElseIf dtgDe.Text = "False" Then
            dtgDe.Text = "缺勤"
            For ii = 0 To 8
                dtgDe.Col = ii
                dtgDe.CellForeColor = &HFF&
            Next
        End If
        dtgDe.Col = 0
        If Val(dtgDe.Text) > 0 Then
            dtgDe.Text = NI
            NI = NI + 1
        Else
            NI = 1
        End If
    Next
ElseIf opt3.Value = True Then
    jj = 1: NN = 1: Fy = 0: PxT = 0
    For oo = 1 To La + 30
        dtgDe.Row = oo
        dtgDe.Col = 0
        If dtgDe.Text = "" Then Exit For
        dtgDe.Col = 3
'''''        If Month(dtgDe.Text) / 3 <= jj Then '累加
             dtgDe.Col = 7: Fy = Fy + Val(dtgDe.Text)
             dtgDe.Col = 5
             If dtgDe.Text = "True" Then
                dtgDe.Col = 6: PxT = PxT + Val(dtgDe.Text)
             End If
             dtgDN.Row = NN
             For ii = 0 To 8
                dtgDe.Col = ii: dtgDN.Col = ii
                dtgDN.Text = dtgDe.Text
             Next
             NN = NN + 1
'''''        Else
'''''            dtgDN.Row = NN
'''''            dtgDN.Col = 1: dtgDN.Text = lblYear & "年度统计:"
'''''            dtgDN.Col = 6: dtgDN.Text = PxT
'''''            dtgDN.Col = 7: dtgDN.Text = Fy
'''''            Fy = 0: PxT = 0
'''''            NN = NN + 1
'''''            jj = jj + 1
'''''            oo = oo - 1
'''''        End If
    Next
    For YY = 1 To 1
            dtgDN.Row = NN
            dtgDN.Col = 2: dtgDN.Text = lblYear & "年度统计:"

                dtgDN.Col = 6
                dtgDN.Text = PxT
            dtgDN.Col = 7: dtgDN.Text = Fy
            NN = NN + 1
            PxT = 0: Fy = 0
    Next
    'dtgDe.Clear
    '回拷
    NI = 1
    For oo = 1 To dtgDN.Rows
        dtgDe.Row = oo: dtgDN.Row = oo
        For ii = 0 To 8
            dtgDe.Col = ii: dtgDN.Col = ii
            dtgDe.Text = dtgDN.Text
        Next
        dtgDe.Col = 0
        If dtgDe.Text = "" Then
            dtgDe.RowHeight(oo) = dtgDe.RowHeight(0)
        Else
            dtgDe.RowHeight(oo) = dtgDe.RowHeight(0) * 3
        End If
        dtgDe.Col = 5
        If dtgDe.Text = "True" Then
            dtgDe.Text = "出勤"
        ElseIf dtgDe.Text = "False" Then
            dtgDe.Text = "缺勤"
            For ii = 0 To 8
                dtgDe.Col = ii
                dtgDe.CellForeColor = &HFF&
            Next
        End If
        dtgDe.Col = 0
        If Val(dtgDe.Text) > 0 Then
            dtgDe.Text = NI
            NI = NI + 1
        Else
            NI = 1
        End If
    Next
End If
dtgDe.Visible = True
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
lblYear.Caption = Year(mod1.DQda)
If mod1.DName = "宋晓炯　" Or mod1.DName = "马晓聪" Or mod1.DName = "吴之禺" Or mod1.DName = "陈珊珊" Then
    cmdAll.Visible = True
Else
    cmdAll.Visible = False
End If
End Sub



Public Sub dtgDeFF()
dtgDe.Clear
dtgDe.Rows = 50
dtgDe.Cols = 9
dtgDe.Row = 0
dtgDe.Col = 0: dtgDe.Text = "序号": dtgDe.CellFontBold = True
dtgDe.Col = 1: dtgDe.Text = "培训种类": dtgDe.CellFontBold = True
dtgDe.Col = 2: dtgDe.Text = "培训课程": dtgDe.CellFontBold = True
dtgDe.Col = 3: dtgDe.Text = "时间": dtgDe.CellFontBold = True
dtgDe.Col = 4: dtgDe.Text = "地点": dtgDe.CellFontBold = True
dtgDe.Col = 5: dtgDe.Text = "出勤": dtgDe.CellFontBold = True
dtgDe.Col = 6: dtgDe.Text = "学时": dtgDe.CellFontBold = True
dtgDe.Col = 7: dtgDe.Text = "费用": dtgDe.CellFontBold = True
dtgDe.Col = 8: dtgDe.Text = "备注": dtgDe.CellFontBold = True
dtgDe.ColWidth(0) = 660
dtgDe.ColWidth(2) = 1500
dtgDe.ColWidth(3) = 960
dtgDe.ColWidth(4) = 1440
dtgDe.ColWidth(5) = 555
dtgDe.ColWidth(6) = 570
End Sub

Public Sub BRBound(Ra)
'select name,cq,cq1,zt,zfy,userZw from peiView3
Dim La As Long
Dim oo As Long
On Error Resume Next
Call Me.dtgbrFF
La = UBound(Ra, 2) + 1
dtgBr.Rows = La + 50
dtgN.Rows = La + 50
For oo = 1 To La
    dtgBr.Row = oo
    dtgBr.Col = 0: dtgBr.Text = Ra(0, oo - 1) '姓名
    dtgBr.Col = 1: dtgBr.Text = Ra(1, oo - 1) '参训次数
    dtgBr.Col = 2: dtgBr.Text = Ra(3, oo - 1) '参训总时数
    dtgBr.Col = 3: dtgBr.Text = Str(Round(Ra(1, oo - 1) / Ra(2, oo - 1), 2) * 100) & "%" '参训出勤率
    dtgBr.Col = 4: dtgBr.Text = Round(Ra(4, oo - 1), 2) '培训总费用
    dtgBr.Col = 5: dtgBr.Text = Ra(5, oo - 1)
    dtgBr.Col = 6: dtgBr.Text = Ra(6, oo - 1)
    
    dtgN.Row = oo
    dtgN.Col = 0: dtgN.Text = Ra(0, oo - 1) '姓名
    dtgN.Col = 1: dtgN.Text = Ra(1, oo - 1) '参训次数
    dtgN.Col = 2: dtgN.Text = Ra(3, oo - 1) '参训总时数
    dtgN.Col = 3: dtgN.Text = Str(Round(Ra(1, oo - 1) / Ra(2, oo - 1), 2) * 100) & "%" '参训出勤率
    dtgN.Col = 4: dtgN.Text = Round(Ra(4, oo - 1), 2) '培训总费用
    dtgN.Col = 5: dtgN.Text = Ra(5, oo - 1)
    dtgN.Col = 6: dtgN.Text = Ra(6, oo - 1)
Next
End Sub

Private Sub opt1_Click()
Call dtgBr_Click
End Sub


Private Sub opt2_Click()
Call dtgBr_Click
End Sub


Private Sub opt3_Click()
Call dtgBr_Click
End Sub

Public Sub BRBMBound(Ra, Rb)
'select name,cq,cq1,zt,zfy,userZw from peiView3
Dim La As Long: Dim Lb As Long
Dim oo As Long
Dim ii As Long
On Error Resume Next
Call Me.dtgbrFF
La = UBound(Ra, 2) + 1
Lb = UBound(Rb, 2) + 1
dtgBr.Rows = La + 50
dtgN.Rows = La + 50
For oo = 1 To Lb
    dtgBr.Row = oo
    dtgBr.Col = 0: dtgBr.Text = Rb(0, oo - 1)
    dtgN.Row = oo
    dtgN.Col = 0: dtgN.Text = Rb(0, oo - 1)
    For ii = 0 To Lb - 1
        If Ra(0, ii) = dtgBr.Text Then
            dtgBr.Col = 1: dtgBr.Text = Ra(1, ii) '参训次数
            dtgBr.Col = 2: dtgBr.Text = Ra(3, ii) '参训总时数
            dtgBr.Col = 3: dtgBr.Text = Str(Round(Ra(1, ii) / Ra(2, ii), 2) * 100) & "%" '参训出勤率
            dtgBr.Col = 4: dtgBr.Text = Round(Ra(4, ii), 2) '培训总费用
            dtgBr.Col = 5: dtgBr.Text = Ra(5, ii)
            dtgBr.Col = 6: dtgBr.Text = Ra(6, ii)
            dtgN.Col = 1: dtgN.Text = Ra(1, ii) '参训次数
            dtgN.Col = 2: dtgN.Text = Ra(3, ii) '参训总时数
            dtgN.Col = 3: dtgN.Text = Str(Round(Ra(1, ii) / Ra(2, ii), 2) * 100) & "%" '参训出勤率
            dtgN.Col = 4: dtgN.Text = Round(Ra(4, ii), 2) '培训总费用
            dtgN.Col = 5: dtgN.Text = Ra(5, ii)
            dtgN.Col = 6: dtgN.Text = Ra(6, ii)
        End If
    Next
Next
'''''For ii = 1 To Lb
'''''    dtgBr.Row = ii
'''''    For oo = 0 To La - 1
'''''
'''''        dtgBr.Col = 0: dtgBr.Text = Ra(0, oo - 1) '姓名
'''''        dtgBr.Col = 1: dtgBr.Text = Ra(1, oo - 1) '参训次数
'''''        dtgBr.Col = 2: dtgBr.Text = Ra(3, oo - 1) '参训总时数
'''''        dtgBr.Col = 3: dtgBr.Text = Str(Round(Ra(1, oo - 1) / Ra(2, oo - 1), 2) * 100) & "%" '参训出勤率
'''''        dtgBr.Col = 4: dtgBr.Text = Round(Ra(4, oo - 1), 2) '培训总费用
'''''        dtgBr.Col = 5: dtgBr.Text = Ra(5, oo - 1)
'''''        dtgBr.Col = 6: dtgBr.Text = Ra(6, oo - 1)
'''''
'''''    Next
'''''Next
End Sub
