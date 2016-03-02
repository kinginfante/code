VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcZuiBrow 
   BackColor       =   &H00FFFFC0&
   Caption         =   "成本追加一览"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdC 
      BackColor       =   &H00C0FFC0&
      Caption         =   "查询"
      Height          =   345
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8490
      Width           =   1005
   End
   Begin VB.TextBox txtZ 
      Height          =   285
      Left            =   900
      TabIndex        =   4
      Top             =   8490
      Width           =   2535
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   405
      Left            =   8940
      TabIndex        =   2
      Top             =   8430
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   714
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14550
      Picture         =   "FmxcZuiBrow.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8340
      Width           =   585
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8205
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   14473
      _Version        =   393216
      BackColor       =   12648384
      BackColorFixed  =   12648384
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "查询:"
      Height          =   285
      Left            =   270
      TabIndex        =   3
      Top             =   8550
      Width           =   705
   End
End
Attribute VB_Name = "FmxcZuiBrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tt As String

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
End Sub

Private Sub cmdC_Click()
Dim ZT As Integer
ZT = 0
If txtZ.Text = "评审通过" Then
    ZT = 100
ElseIf txtZ.Text = "执行" Then
    ZT = 101
ElseIf txtZ.Text = "评审" Then
    ZT = 99
End If

    If mod1.Qy = "上海" Then
        Me.tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc,newF from htzuiView where khmc like '%" & txtZ.Text & "%' or bh like '%" & txtZ.Text & "%'" & _
        " or htxz='" & txtZ.Text & "'  or ywy='" & txtZ.Text & "' order by ztime desc,zid desc"
        If ZT > 0 Then
        Me.tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc,newF from htzuiView where lc=" & ZT & " order by ztime desc,zid desc"
        End If
        If ZT = 99 Then '评审
        Me.tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc,newF from htzuiView where lc<99 order by ztime desc,zid desc"
        End If
        If Val(txtZ.Text) > 0 Then
        Me.tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc,newF from htzuiView where htze=" & Val(txtZ.Text) & " or ze=" & Val(txtZ.Text) & " or bh like '%" & txtZ.Text & "%' order by ztime desc,zid desc"
        End If
    Else
        Me.tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc,newF from htzuiView  where qy='" & mod1.Qy & "'  khmc like '%" & txtZ.Text & "%' or bh like '%" & txtZ.Text & "%'" & _
        " or htxz='" & txtZ.Text & "'  or ywy='" & txtZ.Text & "' order by ztime desc,zid desc"
        If ZT > 0 Then
        Me.tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc,newF from htzuiView  where qy='" & mod1.Qy & "' and lc=" & ZT & " order by ztime desc,zid desc"
        End If
        If ZT = 99 Then '评审
        Me.tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc,newF from htzuiView  where qy='" & mod1.Qy & "'  and lc<99  order by ztime desc,zid desc"
        End If
        If Val(txtZ.Text) > 0 Then
        Me.tt = "select khmc,bh,htze,htxz,ze,ywy,zid,lc,newF from htzuiView where qy='" & mod1.Qy & "' and (htze=" & Val(txtZ.Text) & " or ze=" & Val(txtZ.Text) & ") or bh like '%" & txtZ.Text & "%' order by ztime desc,zid desc"
        End If
    End If
    Call Me.Bound(tt)
End Sub

Private Sub dtgBr_DblClick()
dtgN.Row = dtgBr.Row
dtgN.Col = 6
Call fmxcZJ.Bound(Val(dtgN.Text))
fmxcZJ.Show
fmxcZJ.ZOrder 0

End Sub

Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Call Me.dtgFF
End Sub

Public Sub dtgFF()
dtgBr.Cols = 10
dtgBr.Rows = 30
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "客户名称": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "编号": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "合同金额": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "合同性质": dtgBr.CellFontBold = True
dtgBr.Col = 4: dtgBr.Text = "成本金额": dtgBr.CellFontBold = True
dtgBr.Col = 5: dtgBr.Text = "业务员": dtgBr.CellFontBold = True
dtgBr.Col = 7: dtgBr.Text = "状态": dtgBr.CellFontBold = True
dtgBr.Col = 8: dtgBr.Text = "执行时间": dtgBr.CellFontBold = True
dtgBr.ColWidth(6) = 0
dtgBr.ColWidth(5) = 1080
dtgBr.ColWidth(4) = 1605
dtgBr.ColWidth(3) = 1455
dtgBr.ColWidth(2) = 1515
dtgBr.ColWidth(1) = 2520
dtgBr.ColWidth(0) = 3660
dtgBr.ColWidth(7) = 1000
dtgBr.ColWidth(7) = 2000
dtgBr.ColWidth(9) = 0
dtgN.Cols = 10
dtgN.Rows = 30
dtgN.Row = 0
dtgN.Col = 0: dtgN.Text = "客户名称": dtgN.CellFontBold = True
dtgN.Col = 1: dtgN.Text = "编号": dtgN.CellFontBold = True
dtgN.Col = 2: dtgN.Text = "合同金额": dtgN.CellFontBold = True
dtgN.Col = 3: dtgN.Text = "合同性质": dtgN.CellFontBold = True
dtgN.Col = 4: dtgN.Text = "成本金额": dtgN.CellFontBold = True
dtgN.Col = 5: dtgN.Text = "业务员": dtgN.CellFontBold = True
dtgN.Col = 7: dtgN.Text = "状态": dtgN.CellFontBold = True
dtgN.ColWidth(6) = 0
dtgN.ColWidth(5) = 1080
dtgN.ColWidth(4) = 1605
dtgN.ColWidth(3) = 1455
dtgN.ColWidth(2) = 1515
dtgN.ColWidth(1) = 2720
dtgN.ColWidth(0) = 5460
dtgN.ColWidth(7) = 1000
End Sub

Public Sub Bound(tt As String)
Dim oo As Integer
Dim Ra
Dim La As Integer
dtgBr.Visible = False
Call Me.Qing
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly
On Error Resume Next
Ra = mod1.HTP.GetRows
La = UBound(Ra, 2) + 1
Me.dtgBr.Rows = La + 30
For oo = 1 To La
    dtgBr.Row = oo
    dtgBr.Col = 0: dtgBr.Text = Ra(0, oo - 1)
    dtgBr.Col = 1: dtgBr.Text = Ra(1, oo - 1)
    dtgBr.Col = 2: dtgBr.Text = Ra(2, oo - 1)
    dtgBr.Col = 3: dtgBr.Text = Ra(3, oo - 1)
    dtgBr.Col = 4: dtgBr.Text = Ra(4, oo - 1)
    dtgBr.Col = 5: dtgBr.Text = Ra(5, oo - 1)
    dtgBr.Col = 6: dtgBr.Text = Ra(6, oo - 1)
    dtgBr.Col = 7
    If Ra(7, oo - 1) = 100 Then
        dtgBr.Text = "评审通过"
    ElseIf Ra(7, oo - 1) = 101 Then
        dtgBr.Text = "执行"
    Else
        dtgBr.Text = "评审"
    End If
    dtgBr.Col = 8
    dtgBr.Text = Ra(8, oo - 1)
    dtgN.Row = oo
    dtgN.Col = 0: dtgN.Text = Ra(0, oo - 1)
    dtgN.Col = 1: dtgN.Text = Ra(1, oo - 1)
    dtgN.Col = 2: dtgN.Text = Ra(2, oo - 1)
    dtgN.Col = 3: dtgN.Text = Ra(3, oo - 1)
    dtgN.Col = 4: dtgN.Text = Ra(4, oo - 1)
    dtgN.Col = 5: dtgN.Text = Ra(5, oo - 1)
    dtgN.Col = 6: dtgN.Text = Ra(6, oo - 1)
Next
dtgBr.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
frmZu.Enabled = True
Cancel = True
End Sub

Public Sub Qing()
dtgBr.Clear
dtgN.Clear
Call dtgFF
End Sub

Private Sub txtZ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmdC_Click
End If
End Sub


