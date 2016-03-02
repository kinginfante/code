VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcCG 
   BackColor       =   &H00C0FFC0&
   Caption         =   "新版采购单"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdC 
      Caption         =   "查询"
      Height          =   285
      Left            =   3300
      TabIndex        =   7
      Top             =   8520
      Width           =   1095
   End
   Begin VB.TextBox txtZ 
      Height          =   285
      Left            =   1020
      TabIndex        =   6
      Top             =   8490
      Width           =   2085
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBrN 
      Height          =   585
      Left            =   7140
      TabIndex        =   4
      Top             =   8460
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   1032
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgCDN 
      Height          =   675
      Left            =   10500
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   1191
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   60
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   675
      Left            =   14490
      Picture         =   "FmxcCG.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "返回"
      Top             =   8370
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8265
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   14579
      _Version        =   393216
      BackColor       =   16777152
      Rows            =   50
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   16777152
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgCD 
      Height          =   8265
      Left            =   8670
      TabIndex        =   2
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   14579
      _Version        =   393216
      BackColor       =   12648447
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648447
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "筛选查询"
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   8550
      Width           =   855
   End
End
Attribute VB_Name = "FmxcCG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GyId As Long
Dim Cid As Long
Dim timZm As Integer
Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True

End Sub

Private Sub dtgBr_DblClick()
Dim Cid As Long
dtgBrN.Row = dtgBr.Row
If dtgBrN.Row = 0 Then Exit Sub
dtgBrN.Col = 5
Cid = Val(dtgBrN.Text)
Call FmxcCGDetail.Bound(Cid)
    FmxcCGDetail.Show
    FmxcCGDetail.ZOrder 0
End Sub


Private Sub dtgCD_Click()
dtgCDN.Row = dtgCD.Row
dtgCDN.Col = 5
GyId = Val(dtgCDN.Text)
End Sub

Private Sub dtgCD_DblClick()
Dim ii As Integer
If GyId = 0 Then Exit Sub
ii = MsgBox("是否添加此供应商的采购合同？", vbYesNo + vbQuestion, "请问")

If ii = vbNo Then Exit Sub
timZm = 1 '添加
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "采购合同"
    mod1.cmd.Parameters("@NBLX") = "添加"
    mod1.cmd.Parameters("@bh") = ""
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = GyId
    mod1.cmd.Parameters("@mb1") = Null
    mod1.cmd.Parameters("@md1") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 2 Then '保存
            cmdSave.Enabled = False
        End If
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True



    End If


Set mod1.cmd = Nothing
dtgCD.Enabled = False

End Sub


Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
End Sub

Public Sub dtgBRFF()
dtgBr.Cols = 6
dtgBr.Clear
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "编号": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "供应商": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "采购货品": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "执行状态": dtgBr.CellFontBold = True
dtgBr.Col = 4: dtgBr.Text = "采购日期": dtgBr.CellFontBold = True
dtgBr.ColWidth(1) = 2865
dtgBr.ColWidth(2) = 2370
dtgBrN.Cols = 6
dtgBrN.Clear
dtgBrN.Row = 0:
dtgBrN.Col = 0: dtgBrN.Text = "编号": dtgBrN.CellFontBold = True
dtgBrN.Col = 1: dtgBrN.Text = "供应商": dtgBrN.CellFontBold = True
dtgBrN.Col = 2: dtgBrN.Text = "采购货品": dtgBrN.CellFontBold = True
dtgBrN.Col = 3: dtgBrN.Text = "执行状态": dtgBrN.CellFontBold = True
dtgBrN.Col = 4: dtgBrN.Text = "采购日期": dtgBrN.CellFontBold = True
dtgBr.ColWidth(5) = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
frmZu.Enabled = True
Cancel = True

End Sub

Public Sub dtgCDFF()
dtgCD.Clear
dtgCD.Cols = 6
dtgCD.Row = 0
dtgCD.Col = 0: dtgCD.Text = "货品编号": dtgCD.CellFontBold = True
dtgCD.Col = 1: dtgCD.Text = "数量": dtgCD.CellFontBold = True
dtgCD.Col = 2: dtgCD.Text = "供应商": dtgCD.CellFontBold = True
dtgCD.Col = 3: dtgCD.Text = "合同": dtgCD.CellFontBold = True
dtgCD.Col = 4: dtgCD.Text = "执行日期": dtgCD.CellFontBold = True
dtgCD.ColWidth(0) = 870
dtgCD.ColWidth(1) = 495
dtgCD.ColWidth(2) = 2430
dtgCD.ColWidth(3) = 645
dtgCD.ColWidth(1) = 495
dtgCD.ColWidth(4) = 1770
dtgCD.ColWidth(5) = 0
dtgCDN.Clear
dtgCDN.Cols = 6

End Sub

Public Sub CDBound()
Dim Ra
Dim La As Long
Dim oo As Long
Dim tt As String
tt = "select ljbh,sl,mc,hid,ddrq,gyid from cght where (sl-yl)>0 order by ddrq"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Call dtgCDFF
dtgCD.Rows = La + 50: dtgCDN.Rows = La + 50
For oo = 1 To La
    dtgCD.Row = oo
    dtgCD.Col = 0: dtgCD.Text = Ra(0, oo - 1)
    dtgCD.Col = 1: dtgCD.Text = Ra(1, oo - 1)
    dtgCD.Col = 2: dtgCD.Text = Ra(2, oo - 1)
    dtgCD.Col = 3: dtgCD.Text = Ra(3, oo - 1)
    dtgCD.Col = 4: dtgCD.Text = Ra(4, oo - 1)
    dtgCD.Col = 5: dtgCD.Text = Ra(5, oo - 1)
    dtgCDN.Row = oo
    dtgCDN.Col = 0: dtgCDN.Text = Ra(0, oo - 1)
    dtgCDN.Col = 1: dtgCDN.Text = Ra(1, oo - 1)
    dtgCDN.Col = 2: dtgCDN.Text = Ra(2, oo - 1)
    dtgCDN.Col = 3: dtgCDN.Text = Ra(3, oo - 1)
    dtgCDN.Col = 4: dtgCDN.Text = Ra(4, oo - 1)
    dtgCDN.Col = 5: dtgCDN.Text = Ra(5, oo - 1)
Next

End Sub

Private Sub timQuit_Timer()
Dim Rz
Dim Lz As Integer
Dim Rb
Dim Lb As Integer
Dim RD
Dim Ld As Integer
On Error Resume Next
Dim ii As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '
    dtgCD.Enabled = True
    Call FmxcCGDetail.Bound(Cid)
    FmxcCGDetail.Show
    FmxcCGDetail.ZOrder 0
    Call Me.CDBound
    Call Me.CGBound
End If
timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim Bid As Long
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mm2,mt2,mt1,mt3,mt4 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    mod1.Ti = 5
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    timWait.Enabled = False
    If timZm = 1 Then
        
        Cid = mod1.WP.Fields("mm1").Value
    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    timWait.Enabled = False
    Unload frmWaitA
    Me.Enabled = True
    Exit Sub
'''''    If timZm = 1 Then
'''''        NiceButton1.Enabled = False
'''''    End If
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
        If timZm = 1 Then
            dtgCD.Enabled = True
        End If
    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub



Public Sub CGBound()
Dim tt As String
Dim Ra
Dim La As Integer
Dim oo As Integer
Call dtgBRFF
tt = "select cid,mc,'',lc,drq,cid from cgbound order by cid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgBr.Rows = La + 30
dtgBrN.Rows = dtgBr.Rows

For oo = 1 To La
    dtgBr.Row = oo
    dtgBr.Col = 0: dtgBr.Text = Ra(0, oo - 1)
    dtgBr.Col = 1: dtgBr.Text = Ra(1, oo - 1)
    dtgBr.Col = 2: dtgBr.Text = Ra(2, oo - 1)
    dtgBr.Col = 3: dtgBr.Text = Ra(3, oo - 1)
    dtgBr.Col = 4: dtgBr.Text = Ra(4, oo - 1)
    dtgBr.Col = 5: dtgBr.Text = Ra(5, oo - 1)
    dtgBrN.Row = oo
    dtgBrN.Col = 0: dtgBrN.Text = Ra(0, oo - 1)
    dtgBrN.Col = 1: dtgBrN.Text = Ra(1, oo - 1)
    dtgBrN.Col = 2: dtgBrN.Text = Ra(2, oo - 1)
    dtgBrN.Col = 3: dtgBrN.Text = Ra(3, oo - 1)
    dtgBrN.Col = 4: dtgBrN.Text = Ra(4, oo - 1)
    dtgBrN.Col = 5: dtgBrN.Text = Ra(5, oo - 1)
Next


End Sub
