VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FydED 
   Caption         =   "费用设置"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.OptionButton opt2 
      Caption         =   "修改"
      Height          =   255
      Left            =   9240
      TabIndex        =   15
      Top             =   4440
      Width           =   1815
   End
   Begin VB.OptionButton opt1 
      Caption         =   "修改"
      Height          =   255
      Left            =   7080
      TabIndex        =   14
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Frame frmED 
      Caption         =   "编辑"
      Height          =   2415
      Left            =   7080
      TabIndex        =   7
      Top             =   4800
      Width           =   2055
      Begin VB.CommandButton cmdGB 
         Caption         =   "关    闭"
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "部门添加"
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "所有部门"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox comBm 
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删    除"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton cmdREF 
         Caption         =   "重    置"
         Height          =   375
         Left            =   0
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBm1N 
      Height          =   615
      Left            =   11640
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1085
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLbN 
      Height          =   975
      Left            =   12240
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1720
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBm1 
      Height          =   4455
      Left            =   7080
      TabIndex        =   2
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   20
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLb 
      Height          =   9255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   16325
      _Version        =   393216
      Rows            =   40
      Cols            =   5
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton cmdBack 
      Height          =   375
      Left            =   14760
      Picture         =   "fydED.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "返回"
      Top             =   8760
      Width           =   465
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBm2 
      Height          =   4455
      Left            =   9120
      TabIndex        =   3
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   20
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBm2N 
      Height          =   615
      Left            =   12840
      TabIndex        =   6
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   1085
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FydED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim liD As Long
Dim Bmid1 As Long
Dim Bmid2 As Long
Private Sub cmdAdd_Click()
Dim tt As String
Dim Ra
Dim La
Dim oo As Integer
Dim ii As Integer
On Error GoTo frmED1
If comBm.Text = "" Then
    Exit Sub
End If
If opt1.Value = True Then
    tt = "insert into fydLbBm1 (lid,bmid) values (" & liD & "," & comBm.ItemData(comBm.ListIndex) & ")"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
    Call bm1Bound(liD)
Else
    tt = "insert into fydLbBm2 (lid,bmid) values (" & liD & "," & comBm.ItemData(comBm.ListIndex) & ")"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
    Call bm2Bound(liD)

End If
Exit Sub
frmED1:
MsgBox "出错!"
End Sub

Private Sub cmdAll_Click()
Dim tt As String
Dim Ra
Dim La
Dim oo As Integer
Dim ii As Integer
On Error GoTo frmED4
If opt1.Value = True Then
    tt = "DELETE from fydlbbm1 where lid=" & liD & ";" & _
         "insert into fydLbBm1 (lid,bmid) select " & liD & ",bmid from bm where zzf=1 order by bmid"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
    Call bm1Bound(liD)
Else
    tt = "DELETE from fydlbbm2 where lid=" & liD & ";" & _
         "insert into fydLbBm2 (lid,bmid) select " & liD & ",bmid from bm where zzf=1 order by bmid"
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
    Call bm2Bound(liD)

End If
Exit Sub
frmED4:
MsgBox "出错!"
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub cmdDel_Click()
Dim tt As String
Dim Ra
Dim La
Dim oo As Integer
Dim ii As Integer
On Error GoTo frmED2
If opt1.Value = True Then
    If Bmid1 = 0 Then
        Exit Sub
    End If
    tt = "DELETE from fydlbbm1 where bmid=" & Bmid1 & " and lid=" & liD
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
    Call bm1Bound(liD)
Else
    If Bmid2 = 0 Then
        Exit Sub
    End If
    tt = "DELETE from fydlbbm2 where bmid=" & Bmid2 & " and lid=" & liD
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
    Call bm2Bound(liD)

End If
Exit Sub
frmED2:
MsgBox "出错!"
End Sub


Private Sub cmdGB_Click()
frmED.Visible = False
opt1.Value = False
opt2.Value = False
End Sub

Private Sub cmdRef_Click()
Dim tt As String
Dim Ra
Dim La
Dim oo As Integer
Dim ii As Integer
On Error GoTo frmED3
If opt1.Value = True Then
    tt = "DELETE from fydlbbm1 where lid=" & liD
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
    Call bm1Bound(liD)
Else
    tt = "DELETE from fydlbbm2 where lid=" & liD
    Set mod1.HTP = New ADODB.Recordset
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Set mod1.HTP = Nothing
    Call bm2Bound(liD)

End If
Exit Sub
frmED3:
MsgBox "出错!"
End Sub


Private Sub dtgBm1_Click()
On Error Resume Next
dtgBm1N.Row = dtgBm1.Row
dtgBm1N.Col = 1
Bmid1 = Val(dtgBm1N.Text)
End Sub

Private Sub dtgBm2_Click()
On Error Resume Next
dtgBm2N.Row = dtgBm2.Row
dtgBm2N.Col = 1
Bmid2 = Val(dtgBm2N.Text)
End Sub


Private Sub dtgLb_Click()
dtgLbN.Row = dtgLb.Row
dtgLbN.Col = 4
liD = Val(dtgLbN.Text)
Call bm1Bound(liD)
Call bm2Bound(liD)
End Sub

Private Sub Form_Load()
Dim tt As String
Dim Ra
Dim La
Dim oo As Integer
Dim ii As Integer
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
dtgLb.ColWidth(0) = 300
dtgLb.ColWidth(4) = 0
dtgBm1.ColWidth(0) = 2000
dtgBm1.Row = 0: dtgBm1.Col = 0: dtgBm1.Text = "申请部门": dtgBm1.CellFontBold = True: dtgBm1.Rows = 20
dtgBm2.ColWidth(0) = 2000
dtgBm2.Row = 0: dtgBm2.Col = 0: dtgBm2.Text = "费用归属部门": dtgBm2.CellFontBold = True: dtgBm2.Rows = 20
tt = "select bm,bmid from bm where zzf=1 order by bmid"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2)
For oo = 0 To La
    comBm.AddItem Ra(0, oo), oo
    comBm.ItemData(oo) = Ra(1, oo)
Next


Set mod1.HTP = New ADODB.Recordset
tt = "select * from fydLb order by Lid"
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) - 1
FydED.dtgLb.Rows = La + 10
dtgLbN.Rows = dtgLb.Rows: dtgLbN.Cols = dtgLb.Cols
Call FydED.LBQing
For oo = 1 To La + 2
    FydED.dtgLb.Row = oo
    dtgLbN.Row = oo
    For ii = 1 To 4
        FydED.dtgLb.Col = ii
        dtgLbN.Col = ii
        If IsNull(Ra(ii - 1, oo - 1)) = True Then
            FydED.dtgLb.Text = ""
            dtgLbN.Text = ""
        Else
            FydED.dtgLb.Text = Trim(Ra(ii - 1, oo - 1))
            dtgLbN.Text = Trim(Ra(ii - 1, oo - 1))
        End If
    Next
Next
End Sub

Public Sub LBQing()
dtgLb.Clear
dtgLb.Col = 1: dtgLb.Row = 0: dtgLb.Text = "费用类别": dtgLb.CellFontBold = True
dtgLb.Col = 2: dtgLb.Text = "流转文件": dtgLb.CellFontBold = True
dtgLb.Col = 3: dtgLb.Text = "备注": dtgLb.CellFontBold = True
dtgLb.ColWidth(1) = 1650: dtgLb.ColWidth(2) = 2250: dtgLb.ColWidth(3) = 4250
End Sub

Public Sub bm1Bound(liD As Long)
Dim tt As String
Dim oo As Integer
Dim ii As Integer
Dim Ra
Dim La
On Error Resume Next
dtgBm1.Clear: dtgBm1N.Clear
dtgBm1.Row = 0: dtgBm1.Col = 0: dtgBm1.Text = "申请部门": dtgBm1.CellFontBold = True
dtgBm1.ColWidth(1) = 0
 dtgBm1N.Cols = dtgBm1.Cols

tt = "SELECT dbo.BM.BM, dbo.FydLBbm1.bmid FROM dbo.fydLB INNER JOIN dbo.FydLBbm1 ON dbo.fydLB.Lid = dbo.FydLBbm1.Lid INNER JOIN dbo.BM ON dbo.FydLBbm1.bmid = dbo.BM.BMID" & _
     " where dbo.fydLB.Lid =" & liD & " order by dbo.FydLBbm1.bmid"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2)
For oo = 1 To La + 2
    dtgBm1.Row = oo
    dtgBm1N.Row = oo
    For ii = 0 To 1
        dtgBm1.Col = ii
        dtgBm1N.Col = ii
        dtgBm1.Text = Ra(ii, oo - 1)
        dtgBm1N.Text = Ra(ii, oo - 1)
    Next
Next
dtgBm1.Rows = La + 20: dtgBm1N.Rows = dtgBm1.Rows:
End Sub
Public Sub bm2Bound(liD As Long)
Dim tt As String
Dim oo As Integer
Dim ii As Integer
Dim Ra
Dim La
On Error Resume Next
dtgBm2.Clear: dtgBm2N.Clear
dtgBm2.Row = 0: dtgBm2.Col = 0: dtgBm2.Text = "费用归属部门": dtgBm2.CellFontBold = True
dtgBm2.ColWidth(1) = 0
 dtgBm2N.Cols = dtgBm2.Cols

tt = "SELECT dbo.BM.BM, dbo.FydLBbm2.bmid FROM dbo.fydLB INNER JOIN dbo.FydLBbm2 ON dbo.fydLB.Lid = dbo.FydLBbm2.Lid INNER JOIN dbo.BM ON dbo.FydLBbm2.bmid = dbo.BM.BMID" & _
     " where dbo.fydLB.Lid =" & liD & " order by dbo.FydLBbm2.bmid"
Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2)
For oo = 1 To La + 2
    dtgBm2.Row = oo
    dtgBm2N.Row = oo
    For ii = 0 To 1
        dtgBm2.Col = ii
        dtgBm2N.Col = ii
        dtgBm2.Text = Ra(ii, oo - 1)
        dtgBm2N.Text = Ra(ii, oo - 1)
    Next
Next
dtgBm2.Rows = La + 20: dtgBm2N.Rows = dtgBm2.Rows:
End Sub


Private Sub opt1_Click()
frmED.Left = dtgBm1.Left
frmED.Visible = True
End Sub


Private Sub opt2_Click()
frmED.Left = dtgBm2.Left
frmED.Visible = True
End Sub


