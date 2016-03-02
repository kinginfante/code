VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form fmxcRGB 
   BackColor       =   &H00C0FFC0&
   Caption         =   "人工技术资料"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15060
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   15060
   Begin VB.Frame frmQD 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Frame2"
      Height          =   6615
      Left            =   600
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   12375
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgQD 
         Height          =   8325
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   15075
         _ExtentX        =   26591
         _ExtentY        =   14684
         _Version        =   393216
         BackColor       =   16777088
         Rows            =   14
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   16744576
         BackColorBkg    =   16777152
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
   End
   Begin VB.CommandButton cmdFH 
      Caption         =   "返回"
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   80
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdBG 
      Caption         =   "刷新"
      Height          =   375
      Left            =   13440
      TabIndex        =   9
      Top             =   80
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   8880
      TabIndex        =   5
      Top             =   8400
      Width           =   5535
      Begin VB.OptionButton Option2 
         Caption         =   "人工清单表"
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "工作内容表"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1860
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "机型表"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   1815
      Left            =   7680
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   3201
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.OptionButton opt2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "维修"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   2
      Top             =   8520
      Width           =   1095
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "保养"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   1
      Top             =   8520
      Value           =   -1  'True
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15075
      _ExtentX        =   26591
      _ExtentY        =   14684
      _Version        =   393216
      BackColor       =   16777088
      Rows            =   14
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   16744576
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin VB.Label lblK 
      AutoSize        =   -1  'True
      Caption         =   "111111111111111111"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5160
      TabIndex        =   10
      Top             =   8520
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label lblT 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   8400
      Width           =   5775
   End
End
Attribute VB_Name = "fmxcRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim JX As String  '机型
Dim BB1 As String
Dim XL As Integer
Dim TL As String

Private Sub cmdBG_Click()
Dim tt As String
If Option3.Value = True Then

tt = "select fl,j1,j2,j3,j4,j5,j6,j7,bz,jxdm from NJDM  order by jid "

Call Me.dtgBrBound(tt)
lblT.Caption = "": lblT.ToolTipText = ""
Option1.Enabled = False
End If
End Sub

Private Sub cmdFH_Click()
Dim tt As String
cmdFH.Visible = False
    XL = 1
    tt = "select lb1,bb1 from ngdm where lb like '%" & JX & "%' and bb8='维修' group by lb1,bb1  order by sum(gid) desc"
    Call Me.dtgBrBound1(tt)


End Sub

Private Sub dtgBr_Click()
dtgN.Row = dtgBr.Row
If Option3.Value = True Then
    dtgN.Col = 0: lblT.ToolTipText = dtgN.Text
    dtgN.Col = 1: lblT.Caption = dtgN.Text
    dtgN.Col = 2: JX = dtgN.Text
    Option1.Enabled = True
ElseIf Option1.Value = True And opt2.Value = True And cmdFH.Visible = False Then
    dtgN.Col = 3
    BB1 = dtgN.Text
    dtgN.Col = 0
    TL = dtgN.Text
End If
End Sub

Private Sub dtgBr_DblClick()
Dim tt As String
If Option3.Value = True Then
    tt = "select fl,j1,j2,j3,j4,j5,j6,j7,bz,jxdm from NJDM  where jxdm='" & lblT.ToolTipText & "'"
    Call Me.dtgBrBound(tt)
ElseIf Option1.Value = True And opt2.Value = True And cmdFH.Visible = False Then
    XL = 2
    tt = "select bb3 from ngdm where bb1='" & BB1 & "'  order by gid"
    Call Me.dtgBrBound1(tt)
    cmdFH.Visible = True
End If
End Sub


Private Sub Form_Load()

Dim tt As String
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
tt = "select fl,j1,j2,j3,j4,j5,j6,j7,bz,jxdm from NJDM  order by jid "

Call Me.dtgBrBound(tt)
frmQD.Left = 0: frmQD.Width = Me.Width
frmQD.Top = 0: frmQD.Height = 8325
dtgQD.Left = 0: frmQD.Width = Me.Width
dtgQD.Top = 0: frmQD.Height = 8325
End Sub

Public Sub dtgbrFF()
dtgBr.Clear
dtgBr.FixedRows = 1
dtgBr.Cols = 3
dtgBr.Rows = 200
dtgBr.Row = 0
dtgBr.RowHeight(0) = 500
dtgBr.Col = 1:
dtgBr.Text = "机型": dtgBr.CellFontBold = True
dtgBr.Col = 0: dtgBr.Text = "代码": dtgBr.CellFontBold = True
dtgBr.ColWidth(1) = 13470
dtgBr.ColWidth(2) = 0
dtgBr.ColWidth(0) = -1

dtgN.Clear
dtgN.Cols = 3
dtgN.Rows = 200
dtgN.Row = 0
For oo = 1 To 199
    dtgBr.RowHeight(oo) = 315
Next

End Sub


Public Sub dtgbrFF1()
dtgBr.Clear
dtgBr.FixedRows = 1
dtgBr.Cols = 4
dtgBr.Rows = 200
dtgBr.Row = 0
dtgBr.RowHeight(0) = 500
dtgBr.Col = 0:
dtgBr.Text = "类型（维修）": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "豪曼编码": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "工作内容": dtgBr.CellFontBold = True
dtgBr.ColWidth(0) = 3510
dtgBr.ColWidth(1) = 1275
dtgBr.ColWidth(2) = 9930
dtgBr.ColWidth(3) = 0

dtgN.Clear
dtgN.Cols = 4
dtgN.Rows = 200
For oo = 1 To 199
    dtgBr.RowHeight(oo) = 629
Next
End Sub

Public Sub dtgBrBound(tt As String)
Dim Ra
Dim LT As String
Dim La As Integer
dtgBr.Visible = False
Call Me.dtgbrFF
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgBr.Rows = La + 20
dtgN.Rows = La + 20
For oo = 1 To La
    LT = Ra(0, oo - 1)
    If IsNull(Ra(1, oo - 1)) = False Then
        LT = LT & " " & Ra(1, oo - 1)
    End If
    If IsNull(Ra(2, oo - 1)) = False Then
        LT = LT & " " & Ra(2, oo - 1)
    End If
    If IsNull(Ra(3, oo - 1)) = False Then
        LT = LT & " " & Ra(3, oo - 1)
    End If
    If IsNull(Ra(4, oo - 1)) = False Then
        LT = LT & " " & Ra(4, oo - 1)
    End If
    If IsNull(Ra(5, oo - 1)) = False Then
        LT = LT & " " & Ra(5, oo - 1)
    End If
    If IsNull(Ra(6, oo - 1)) = False Then
        LT = LT & " " & Ra(6, oo - 1)
    End If
    If IsNull(Ra(7, oo - 1)) = False Then
        LT = LT & " " & Ra(7, oo - 1)
    End If
    If IsNull(Ra(8, oo - 1)) = False Then
        LT = LT & " " & Ra(8, oo - 1)
    End If
    dtgBr.Row = oo
    dtgBr.Col = 1: dtgBr.Text = LT
    dtgBr.Col = 0: dtgBr.Text = Ra(9, oo - 1)
    dtgN.Row = oo
    dtgN.Col = 1: dtgN.Text = LT
    dtgN.Col = 0: dtgN.Text = Ra(9, oo - 1)
    dtgBr.Col = 2: dtgN.Col = 2
    dtgBr.Text = Ra(0, oo - 1)
    dtgN.Text = Ra(0, oo - 1)
Next
dtgBr.Visible = True
End Sub
Public Sub dtgBrBound1(tt As String)
Dim Ra
Dim LT As String
Dim La As Integer
Call Me.dtgbrFF1
dtgBr.Visible = False
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next
La = UBound(Ra, 2) + 1
dtgBr.Rows = La + 20
dtgN.Rows = La + 20
If opt1.Value = True Then  '维保
    For oo = 1 To La
        dtgBr.Row = oo
        dtgBr.Col = 2: dtgBr.Text = Ra(1, oo - 1)
        dtgBr.Col = 3: dtgBr.Text = Ra(2, oo - 1)
        dtgN.Row = oo
        dtgN.Col = 2: dtgN.Text = Ra(1, oo - 1)
        dtgN.Col = 3: dtgN.Text = Ra(2, oo - 1)
    Next
ElseIf opt2.Value = True And XL = 1 Then '绑定维修分类
    For oo = 1 To La
        dtgBr.Row = oo
        If oo = 28 Then
            oo = oo
        End If
        dtgBr.Col = 0: dtgBr.Text = Ra(0, oo - 1)
        '调整行高
        lblK.Caption = dtgBr.Text
        lblK.AutoSize = True
        If lblK.Width / dtgBr.ColWidth(0) >= 2 Then
            dtgBr.RowHeight(oo) = 315 * (lblK.Width / dtgBr.ColWidth(0) + 1)
        End If
        dtgBr.Col = 3: dtgBr.Text = Ra(1, oo - 1)
        dtgN.Row = oo
        dtgN.Col = 0: dtgN.Text = Ra(0, oo - 1)
        dtgN.Col = 3: dtgN.Text = Ra(1, oo - 1)
    Next
ElseIf opt2.Value = True And XL = 2 Then '绑定维修明细
    For oo = 1 To La
        dtgBr.Row = oo
        dtgBr.Col = 0: dtgBr.Text = ""
        dtgBr.Col = 2: dtgBr.Text = Ra(0, oo - 1)
        '调整行高
        lblK.Caption = dtgBr.Text
        lblK.AutoSize = True
        If lblK.Width / dtgBr.ColWidth(2) >= 2 Then
            dtgBr.RowHeight(oo) = 315 * (lblK.Width / dtgBr.ColWidth(2) + 1)
        End If
        'dtgBr.Col = 3: dtgBr.Text = Ra(1, oo - 1)
        dtgN.Row = oo
        dtgN.Col = 0: dtgN.Text = ""
        dtgN.Col = 2: dtgN.Text = Ra(2, oo - 1)
        'dtgN.Col = 3: dtgN.Text = Ra(1, oo - 1)
    Next
    dtgBr.Row = 1: dtgBr.Col = 0: dtgBr.Text = TL
    dtgN.Row = 1: dtgN.Col = 0: dtgN.Text = TL
End If
dtgBr.Visible = True
End Sub

Private Sub Option1_Click()
Dim tt As String
Dim xZ As String
frmQD.Visible = False
If lblT.Caption = "" Then Exit Sub
XL = 1
Call Me.dtgbrFF1

dtgBr.Row = 0: dtgN.Row = 0
dtgBr.Col = 0: dtgN.Col = 0
If opt1.Value = True Then
    xZ = "保养"
Else
    xZ = "维修"
End If
'''''dtgBr.CellFontBold = True
'''''dtgBr.Text = lblT.Caption & " " & xZ
'''''dtgN.Text = lblT.Caption & " " & xZ
'''''dtgBr.Col = 1: dtgBr.Text = lblT.Caption & " " & xZ
'''''dtgBr.Col = 2: dtgBr.Text = lblT.Caption & " " & xZ

'''''dtgBr.MergeCells = flexMergeFree
'''''dtgBr.MergeCol(0) = True
'''''dtgBr.MergeCol(1) = True
'''''dtgBr.MergeCol(2) = True
'''''dtgBr.MergeRow(0) = True
'''dtgBr.CellAlignment = 2
If opt1.Value = True Then
    tt = "select lb1,bb3,gid from ngdm where lb like '%" & JX & "%' and bb8='保养'  order by gid"
    Call Me.dtgBrBound1(tt)
Else
    tt = "select lb1,bb1 from ngdm where lb like '%" & JX & "%' and bb8='维修' group by lb1,bb1  order by sum(gid) desc"
    Call Me.dtgBrBound1(tt)
    'cmdFH.Visible = True
End If
cmdBG.Visible = False
End Sub


Private Sub Option2_Click()
frmQD.Visible = True
End Sub

Private Sub Option3_Click()
Dim tt As String
frmQD.Visible = False
tt = "select fl,j1,j2,j3,j4,j5,j6,j7,bz,jxdm from NJDM  where jxdm='" & lblT.ToolTipText & "'"
Call Me.dtgBrBound(tt)
cmdFH.Visible = False
cmdBG.Visible = True
XL = 0
End Sub

Public Sub dtgQDFF()
dtgQD.Rows = 500
dtgQD.Cols = 5
dtgQD.Row = 0
dtgQD.Col = 2: dtgQD.Text = "类型": dtgQD.CellFontBold = True
dtgQD.Col = 3: dtgQD.Text = "人工内容": dtgQD.CellFontBold = True
End Sub
