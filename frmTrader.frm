VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmTrader 
   BackColor       =   &H00BFFFE2&
   Caption         =   "供应商列表"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   7380
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox txtZ 
      BackColor       =   &H00BFFFE2&
      Height          =   300
      Left            =   2940
      TabIndex        =   6
      Top             =   5700
      Width           =   2535
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   5790
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   661
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdCA 
      BackColor       =   &H00FFFFC0&
      Caption         =   "查  询"
      Height          =   285
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5730
      Width           =   1125
   End
   Begin VB.ComboBox comLx 
      BackColor       =   &H00BFFFE2&
      Height          =   300
      ItemData        =   "frmTrader.frx":0000
      Left            =   810
      List            =   "frmTrader.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5700
      Width           =   1605
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   5595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   9869
      _Version        =   393216
      BackColor       =   12582882
      Rows            =   25
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   15728356
      BackColorBkg    =   12582882
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "值"
      Height          =   225
      Left            =   2610
      TabIndex        =   3
      Top             =   5760
      Width           =   285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "类型:"
      Height          =   255
      Left            =   150
      TabIndex        =   1
      Top             =   5760
      Width           =   615
   End
End
Attribute VB_Name = "frmTrader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Qy
Public Qa As Integer
Private Sub cmdCA_Click()
Dim tt As String
Dim Ra
Dim La As Long
tt = "SELECT  SD30301_豪曼制冷.dbo.l_trader.code, SD30301_豪曼制冷.dbo.l_area.name as area, SD30301_豪曼制冷.dbo.l_trader.name, SD30301_豪曼制冷.dbo.l_trader.traderid" & _
    " FROM SD30301_豪曼制冷.dbo.l_trader INNER JOIN SD30301_豪曼制冷.dbo.l_area ON SD30301_豪曼制冷.dbo.l_trader.areaid = SD30301_豪曼制冷.dbo.l_area.areaid"
Select Case comLx.Text
Case "编码"
    tt = tt & " where SD30301_豪曼制冷.dbo.l_trader.code like '%" & txtZ.Text & "%'"
Case "供应商名称"
    tt = tt & " where SD30301_豪曼制冷.dbo.l_trader.name like '%" & txtZ.Text & "%'"
Case "区域"
    tt = tt & " where SD30301_豪曼制冷.dbo.l_area.name ='" & txtZ.Text & "'"
End Select

Set mod1.HTP = CreateObject("adodb.recordset")
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Call Bound(Ra, La)
End Sub

Private Sub comLx_Click()
Dim oo As Integer
On Error Resume Next
For oo = txtZ.ListCount To 0 Step -1
    txtZ.RemoveItem oo
Next
If comLx.Text = "区域" Then
    For oo = 0 To Qa
        txtZ.AddItem Qy(0, oo), oo
    Next

End If
End Sub


Private Sub dtgBr_DblClick()
Dim code As String
Dim TraderId As Long
Dim TraderName As String
dtgN.Row = dtgBr.Row
dtgN.Col = 0: code = dtgN.Text
dtgN.Col = 2: TraderName = dtgN.Text
dtgN.Col = 3: TraderId = Val(dtgN.Text)
Me.Visible = False
frmHtz1.txtTrader.Text = TraderName
frmHtz1.txtTrader.ToolTipText = code
frmHtz1.txtTrader.Tag = TraderId
'frmHtz1.ZOrder 0
End Sub

Private Sub Form_Load()
Dim tt As String
dtgBr.Row = 0
dtgBr.ColWidth(2) = 4980
dtgBr.ColWidth(3) = 0
dtgBr.Col = 0: dtgBr.Text = "编码": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "区域": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "供应商": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "ID"
tt = "select SD30301_豪曼制冷.dbo.l_area.name from SD30301_豪曼制冷.dbo.l_area order by SD30301_豪曼制冷.dbo.l_area.areaid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Qy = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Qa = UBound(Qy, 2) + 1
If Qa = 0 Then
    MsgBox "出错!"
    End
End If
End Sub

Public Sub Bound(Ra, La As Long)
Dim oo As Long
Dim ii As Long
On Error Resume Next
dtgBr.Rows = La + 1
dtgN.Rows = dtgBr.Rows
dtgN.Cols = dtgBr.Cols
For oo = 1 To La + 1
    dtgBr.Row = oo: dtgN.Row = oo
    For ii = 0 To 3
        dtgBr.Col = ii: dtgN.Col = ii
        dtgBr.Text = Ra(ii, oo - 1)
        dtgN.Text = Ra(ii, oo - 1)
    Next
Next
End Sub
