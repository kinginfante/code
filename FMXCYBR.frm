VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form FMXCYBR 
   BackColor       =   &H00C0FFC0&
   Caption         =   "到帐情况列表"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制"
      Height          =   315
      Left            =   5670
      TabIndex        =   15
      ToolTipText     =   "点击后，打开Excel，可进行粘贴"
      Top             =   8430
      Width           =   975
   End
   Begin VB.TextBox txtDay 
      Height          =   285
      Left            =   10860
      TabIndex        =   13
      Top             =   8400
      Width           =   315
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   9330
      TabIndex        =   10
      Text            =   "2013"
      Top             =   8400
      Width           =   465
   End
   Begin VB.CommandButton cmdSerach 
      Caption         =   "查询"
      Height          =   315
      Left            =   4410
      TabIndex        =   9
      Top             =   8430
      Width           =   1065
   End
   Begin VB.TextBox txtZ 
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2610
      TabIndex        =   8
      Top             =   8400
      Width           =   1695
   End
   Begin VB.ComboBox comTj 
      Height          =   300
      ItemData        =   "FMXCYBR.frx":0000
      Left            =   1380
      List            =   "FMXCYBR.frx":0013
      TabIndex        =   7
      Text            =   "客户名称"
      Top             =   8400
      Width           =   1245
   End
   Begin VB.CommandButton cmdMon 
      Caption         =   "日期查询"
      Height          =   315
      Left            =   11520
      TabIndex        =   4
      Top             =   8400
      Width           =   885
   End
   Begin VB.TextBox txtMOn 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10200
      TabIndex        =   3
      Top             =   8400
      Width           =   345
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   315
      Left            =   2910
      TabIndex        =   2
      Top             =   8790
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   556
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin NiceFormControl.NiceForm NF 
      Left            =   930
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14400
      Picture         =   "FMXCYBR.frx":0047
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8220
      Width           =   585
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   14314
      _Version        =   393216
      BackColor       =   12648384
      BackColorFixed  =   12648384
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label4 
      Caption         =   "日"
      Height          =   225
      Left            =   11250
      TabIndex        =   14
      Top             =   8430
      Width           =   195
   End
   Begin VB.Label Label3 
      Caption         =   "月"
      Height          =   195
      Left            =   10620
      TabIndex        =   12
      Top             =   8430
      Width           =   195
   End
   Begin VB.Label Label2 
      Caption         =   "年"
      Height          =   195
      Left            =   9870
      TabIndex        =   11
      Top             =   8430
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "查询条件"
      Height          =   285
      Left            =   360
      TabIndex        =   6
      Top             =   8430
      Width           =   1005
   End
   Begin VB.Label lblZE 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   12450
      TabIndex        =   5
      Top             =   8460
      Width           =   1545
   End
End
Attribute VB_Name = "FMXCYBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tt As String

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
End Sub

Private Sub cmdCopy_Click()
If Not (mod1.DName = "顾珣" Or mod1.DName = "乔继敏" Or mod1.DName = "张文琴" Or mod1.DName = "于晓静" Or mod1.DName = "张萍" Or mod1.DName = "陈文超") Then
    Exit Sub
End If
dtgBr.FixedCols = 0
dtgBr.FixedRows = 0
dtgBr.Row = 0
dtgBr.Col = 0
dtgBr.ColSel = 6
dtgBr.RowSel = dtgBr.Rows - 1
Clipboard.Clear
Clipboard.SetText dtgBr.Clip
dtgBr.FixedCols = 1
dtgBr.FixedRows = 1
End Sub

Private Sub cmdMon_Click()
Dim oo As Integer
If Val(txtYear.Text) = 0 Then
    MsgBox "请输入年份"
    Exit Sub
End If
If Val(txtMOn.Text) = 0 And Val(txtDay.Text) > 0 Then
    MsgBox "请输入月份！"
    Exit Sub
End If
If Val(txtYear.Text) > 0 And Val(txtMOn.Text) > 0 And Val(txtDay.Text) > 0 Then
    FMXCYBR.tt = "select khmc,dzrq,je,bz,lc,lcren,aid,htbh,xywy from htAView where year(dzrq)=" & Val(txtYear.Text) & _
    " and month(dzrq)=" & txtMOn.Text & " and day(dzrq)=" & txtDay.Text & " order by aid desc"
ElseIf Val(txtYear.Text) > 0 And Val(txtMOn.Text) > 0 And Val(txtDay.Text) = 0 Then
    FMXCYBR.tt = "select khmc,dzrq,je,bz,lc,lcren,aid,htbh,xywy from htAView where year(dzrq)=" & Val(txtYear.Text) & _
    " and month(dzrq)=" & txtMOn.Text & " order by aid desc"
ElseIf Val(txtYear.Text) > 0 And Val(txtMOn.Text) = 0 And Val(txtDay.Text) = 0 Then
    FMXCYBR.tt = "select khmc,dzrq,je,bz,lc,lcren,aid,htbh,xywy from htAView where year(dzrq)=" & Val(txtYear.Text) & " order by aid desc"
End If

    Call FMXCYBR.REF(FMXCYBR.tt)
    dtgN.Col = 2
For oo = 1 To dtgBr.Rows - 1
    dtgN.Row = oo
    lblZE.Caption = Val(lblZE.Caption) + Val(dtgN.Text)
Next
End Sub

Private Sub Command1_Click()

End Sub


Private Sub cmdSerach_Click()
Dim oo As Integer
Dim companyId As Integer
Select Case comTj.Text
Case "客户名称"

    FMXCYBR.tt = "select khmc,dzrq,je,bz,lc,lcren,aid,htbh,xywy,mbf from htAView where khmc like '%" & txtZ.Text & "%' order by aid desc"
    Call FMXCYBR.REF(FMXCYBR.tt)
Case "业务员"
    FMXCYBR.tt = "select khmc,dzrq,je,bz,lc,lcren,aid,htbh,xywy,mbf from htAView where xywy ='" & txtZ.Text & "' order by aid desc"

Case "收款公司"
    If InStr(1, "豪曼", txtZ.Text) > 0 Then
        companyId = 1
    ElseIf InStr(1, "鼎力", txtZ.Text) > 0 Then
        companyId = 2
    ElseIf InStr(1, "杰升", txtZ.Text) > 0 Then
        companyId = 3
    End If
    FMXCYBR.tt = "select khmc,dzrq,je,bz,lc,lcren,aid,htbh,xywy,mbf from htAView where companyId =" & companyId & " order by aid desc"
Case "到帐金额"
    FMXCYBR.tt = "select khmc,dzrq,je,bz,lc,lcren,aid,htbh,xywy,mbf from htAView where JE =" & Val(txtZ.Text) & " order by aid desc"
Case "合同编号"
    If Left(txtZ.Text, 2) = "HM" Then
        txtZ.Text = Right(txtZ.Text, 5)
    End If
    FMXCYBR.tt = "select khmc,dzrq,je,bz,lc,lcren,aid,htbh,xywy,mbf from htAView where htbh =" & txtZ.Text & " order by aid desc"
    If txtZ.Text = "" Then
        FMXCYBR.tt = "select khmc,dzrq,je,bz,lc,lcren,aid,htbh,xywy,mbf from htAView where htbh is Null order by aid desc"
    End If
End Select
Call FMXCYBR.REF(FMXCYBR.tt)
lblZE.Caption = ""
End Sub


Private Sub dtgBr_DblClick()
dtgN.Row = dtgBr.Row
dtgN.Col = 6
'If Val(dtgN.Text) = 0 Then Exit Sub
On Error Resume Next
Call fmxcY.Bound(Val(dtgN.Text))
fmxcY.Show
fmxcY.ZOrder 0
End Sub

Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
NF.LoadSkin 4
NF.AutoSkinControl
txtMOn.Text = Month(mod1.DQda)
txtYear.Text = Year(mod1.DQda)
txtDay.Text = Day(mod1.DQda)
End Sub

Public Sub REF(tt As String)
Dim Ra
Dim La
Dim ii As Integer
Dim oo As Integer
Dim Oid As Long
Call dtgFF
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgBr.Rows = La + 30: dtgN.Rows = La + 30
dtgBr.Visible = False
For oo = 1 To La
    dtgBr.Row = oo: dtgN.Row = oo
    dtgBr.Col = 0: dtgBr.Text = Replace(Ra(0, oo - 1), Chr(13) + Chr(10), ""): dtgN.Col = 0: dtgN.Text = Ra(0, oo - 1)
    dtgBr.Col = 1: dtgBr.Text = Replace(Ra(1, oo - 1), Chr(13) + Chr(10), ""): dtgN.Col = 1: dtgN.Text = Ra(1, oo - 1)
    dtgBr.Col = 2: dtgBr.Text = Replace(Ra(2, oo - 1), Chr(13) + Chr(10), ""): dtgN.Col = 2: dtgN.Text = Ra(2, oo - 1)
    dtgBr.Col = 3: dtgBr.Text = Replace(Ra(3, oo - 1), Chr(13) + Chr(10), ""): dtgN.Col = 3: dtgN.Text = Ra(3, oo - 1)
    dtgBr.Col = 4: dtgBr.Text = Replace(Ra(4, oo - 1), Chr(13) + Chr(10), ""): dtgN.Col = 4: dtgN.Text = Ra(4, oo - 1)
    dtgBr.Col = 5: dtgBr.Text = Replace(Ra(5, oo - 1), Chr(13) + Chr(10), ""): dtgN.Col = 5: dtgN.Text = Ra(5, oo - 1)
    dtgBr.Col = 6: dtgBr.Text = Replace(Ra(6, oo - 1), Chr(13) + Chr(10), ""): dtgN.Col = 6: dtgN.Text = Ra(6, oo - 1)
    dtgBr.Col = 7: dtgBr.Text = Replace(Ra(9, oo - 1), Chr(13) + Chr(10), ""): dtgN.Col = 7: dtgN.Text = Ra(9, oo - 1)
''''    dtgBr.Col = 6
''''    If dtgBr.Text = "2993" Then
''''    dtgBr.Col = 7
    If dtgBr.Text = False Then
        For ii = 0 To 7
            dtgBr.Col = ii: dtgBr.CellForeColor = &H0&
        Next
    Else
        For ii = 0 To 7
            dtgBr.Col = ii: dtgBr.CellForeColor = &HFF0000
        Next
    End If
'''''    End If
    If IsNull(Ra(7, oo - 1)) = False Then '将合同编号放入备注字体
        dtgBr.Col = 3
        If InStr(0, Ra(7, oo - 1), dtgBr.Text) = 0 Then
            dtgBr.Text = Replace(Trim(dtgBr.Text), Chr(13) + Chr(10), "") & " " & Replace(Ra(7, oo - 1), Chr(13) + Chr(10), "")
        End If
    End If
    dtgBr.Col = 6
    If Val(dtgBr.Text) <> Oid Then
        Oid = Val(dtgBr.Text)
    Else
        dtgBr.Col = 0: dtgBr.Text = ""
        dtgBr.Col = 1: dtgBr.Text = ""
        dtgBr.Col = 2: dtgBr.Text = ""
        dtgBr.Col = 4: dtgBr.Text = ""
        dtgBr.Col = 5: dtgBr.Text = ""
        dtgBr.Col = 6: dtgBr.Text = ""

    End If
Next
dtgBr.Visible = True
End Sub

Public Sub dtgFF()
dtgBr.Rows = 100
dtgBr.Cols = 8
dtgBr.Clear
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.CellFontBold = True: dtgBr.Text = "客户名称"
dtgBr.Col = 1: dtgBr.CellFontBold = True: dtgBr.Text = "到帐日期"
dtgBr.Col = 2: dtgBr.CellFontBold = True: dtgBr.Text = "到帐金额"
dtgBr.Col = 3: dtgBr.CellFontBold = True: dtgBr.Text = "对应合同编号"
dtgBr.Col = 4: dtgBr.CellFontBold = True: dtgBr.Text = "执行状态"
dtgBr.Col = 5: dtgBr.CellFontBold = True: dtgBr.Text = "流程执行"
dtgBr.Col = 6: dtgBr.CellFontBold = True: dtgBr.Text = "ID单号"
dtgBr.Col = 7: dtgBr.CellFontBold = True: dtgBr.Text = "马老师反驳"
dtgBr.ColWidth(0) = 4000
dtgBr.ColWidth(1) = 2000
dtgBr.ColWidth(2) = 1000
dtgBr.ColWidth(3) = 4605
dtgBr.ColWidth(7) = 0
dtgN.Rows = 100
dtgN.Cols = 8
dtgN.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Visible = False
frmZu.Enabled = True
End Sub


