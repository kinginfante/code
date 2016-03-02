VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmComDetail 
   BackColor       =   &H00C0FFC0&
   Caption         =   "计算机配件详情"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15180
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15180
   Begin VB.CommandButton cmdGui 
      BackColor       =   &H00FFFFC0&
      Caption         =   "归档"
      Height          =   585
      Left            =   12270
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8490
      Width           =   675
   End
   Begin VB.CommandButton cmdKZ 
      Caption         =   "视角"
      Height          =   585
      Left            =   12990
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8490
      Width           =   675
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制"
      Height          =   585
      Left            =   13695
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "点击后,打开EXCEL,可将表格内容粘贴."
      Top             =   8490
      Width           =   675
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   585
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8490
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMa 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   14420
      _Version        =   393216
      BackColor       =   16777152
      Rows            =   8
      Cols            =   3
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
End
Attribute VB_Name = "frmComDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZF As Boolean '视图
Public Cid As Integer
Public Sub Initialize()
dtgMa.Clear
dtgMa.Row = 0
dtgMa.Col = 0: dtgMa.Text = "硬件类型": dtgMa.CellFontBold = True: dtgMa.CellAlignment = 0
dtgMa.Col = 1: dtgMa.Text = "标准配置": dtgMa.CellFontBold = True
dtgMa.Col = 2: dtgMa.Text = "改动配置": dtgMa.CellFontBold = True
dtgMa.Row = 1: dtgMa.Col = 0: dtgMa.Text = "芯片": dtgMa.CellFontBold = True
dtgMa.Row = 2: dtgMa.Text = "内存": dtgMa.CellFontBold = True
dtgMa.Row = 3: dtgMa.Text = "硬盘": dtgMa.CellFontBold = True
dtgMa.Row = 4: dtgMa.Text = "主板": dtgMa.CellFontBold = True
dtgMa.Row = 5: dtgMa.Text = "显示器": dtgMa.CellFontBold = True
dtgMa.Row = 6: dtgMa.Text = "光驱": dtgMa.CellFontBold = True
dtgMa.Row = 7: dtgMa.Text = "备注": dtgMa.CellFontBold = True
dtgMa.ColWidth(0) = 1575
dtgMa.ColWidth(1) = 6075
dtgMa.ColWidth(2) = 7185
dtgMa.RowHeight(0) = 405
ZF = False
Call RowH(ZF)
End Sub

Public Sub Bound(Cid As Integer)
Dim tt As String
Dim Ra
Dim Rb
Dim Rc
Dim Lchar As String
Dim oo As Integer
tt = "select uptime,cpudetail,memory+' '+memorydetail,hddetail,mbdetail,monitordetail,cdromdetail,bz from computer where cid=" & Cid & ";" & _
    "select uptime,cpudetail,memory+' '+memorydetail,hddetail,mbdetail,monitordetail,cdromdetail,bz f from computerO where cid=" & Cid & ";" & _
    "select cip from computerIP where cid=" & Cid
    Set mod1.HTP = New ADODB.Recordset
mod1.HTP.Open tt, mod1.wzcc, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rc = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
On Error Resume Next
dtgMa.Visible = False
dtgMa.Col = 1
For oo = 1 To 7
    dtgMa.Row = oo
    dtgMa.Text = Rb(oo, 0)
    dtgMa.CellAlignment = 0
Next

For oo = 1 To 7
    dtgMa.Col = 2
    dtgMa.Row = oo
    dtgMa.Text = Ra(oo, 0)
    dtgMa.CellAlignment = 0
    '比较是否与原配置一致
    Lchar = dtgMa.Text
    dtgMa.Col = 1
    If Trim(Lchar) <> Trim(dtgMa.Text) Then
        'dtgMa.Col = 0: dtgMa.CellForeColor = 255
        dtgMa.Col = 2: dtgMa.CellForeColor = 255
    End If
Next
dtgMa.Row = 0: dtgMa.Col = 0
dtgMa.Text = Rc(0, 0) & Chr(13) & Chr(10) & dtgMa.Text
Me.Cid = Cid
dtgMa.Visible = True
End Sub

Private Sub cmdBack_Click()

Me.Visible = False
End Sub

Private Sub cmdGui_Click()
Dim tt As String
Dim ii As Integer
ii = MsgBox("是否归档?", vbYesNo + vbQuestion, "请注意!")
If ii = vbYes Then
    tt = "Declare @cpu nvarchar(50),@cpudetail nvarchar(3000),@memory nvarchar(50),@memorydetail nvarchar(3000),@hd nvarchar(50),@hddetail nvarchar(3000),@motherboard nvarchar(50),@mbdetail nvarchar(3000)," & _
        "@monitor nvarchar(50),@monitordetail nvarchar(3000),@cdrom nvarchar(50),@cdromdetail nvarchar(3000),@bz nvarchar(3000);" & _
        "select @cpu=cpu,@cpudetail=cpudetail,@memory=memory,@memorydetail=memorydetail,@hd=hd,@hddetail=hddetail,@motherboard=motherboard,@mbdetail=mbdetail," & _
        "@monitor=monitor,@monitordetail=monitordetail,@cdrom=cdrom,@cdromdetail=cdromdetail,@bz=bz from computer where cid=" & Cid & ";" & _
        "update computerO set cpu=@cpu,cpudetail=@cpudetail,memory=@memory,memorydetail=@memorydetail,hd=@hd,hddetail=@hddetail,motherboard=@motherboard,mbdetail=@mbdetail," & _
        "monitor=@monitor,monitordetail=@monitordetail,cdrom=@cdrom,cdromdetail=@cdromdetail,bz=@bz where cid=" & Cid
        Set mod1.HTP = New ADODB.Recordset
        mod1.HTP.Open tt, mod1.wzcc, adOpenForwardOnly, adLockReadOnly, adCmdText
        Set mod1.HTP = Nothing
        Call Me.Initialize
        Call Me.Bound(Cid)
        mod1.Light(Cid) = False
        frmComputer.cmdComputer(Cid).BackColor = frmComputer.cmdComputer(Cid).Tag
End If
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
cmdGui.Visible = False
If mod1.DName = "马晓聪" Then
    cmdGui.Visible = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Cancel = True
End Sub

Public Sub RowH(ZanF As Boolean)
If ZanF = False Then
    dtgMa.RowHeight(7) = 690
    dtgMa.RowHeight(6) = 1665
    dtgMa.RowHeight(5) = 1245
    dtgMa.RowHeight(4) = 1365
    dtgMa.RowHeight(3) = 750
    dtgMa.RowHeight(2) = 990
    dtgMa.RowHeight(1) = 1005
    ZF = True
Else
    ZF = False
End If
End Sub
