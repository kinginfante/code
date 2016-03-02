VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmGxbjL 
   BackColor       =   &H00C0FFC0&
   Caption         =   "零配件查询"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdAgain 
      BackColor       =   &H00C0FFFF&
      Caption         =   "再次搜索"
      Height          =   315
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8700
      Width           =   915
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timQuit 
      Interval        =   1000
      Left            =   660
      Top             =   90
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   465
      Left            =   12360
      TabIndex        =   5
      Top             =   8640
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   820
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   14610
      Picture         =   "frmGxbjL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "返回"
      Top             =   8640
      Width           =   585
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H00C0FFC0&
      Caption         =   "新搜索"
      Height          =   315
      Left            =   3270
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8700
      Width           =   1035
   End
   Begin VB.TextBox txtZ 
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   8700
      Width           =   2925
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8505
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   15002
      _Version        =   393216
      BackColor       =   16777152
      FixedCols       =   0
      BackColorFixed  =   15728356
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
   Begin VB.Label lblTx 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   5850
      TabIndex        =   6
      Top             =   8730
      Width           =   8415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   8730
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "frmGxbjL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timZm As Integer '数据提交后,由timWait执行的后续命令ID(1市场营销部配件添加)
Public tt As String

Private Sub Command1_Click()

End Sub

Private Sub cmdAgain_Click()
Dim ii As Integer
Dim oo As Integer
Dim kk As Integer
dtgBr.Visible = False
dtgBr.Clear
dtgBr.Rows = 30
dtgBr.Cols = 11
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.CellFontBold = True: dtgBr.Text = "品牌"
dtgBr.Col = 1: dtgBr.CellFontBold = True: dtgBr.Text = "编号"
dtgBr.Col = 2: dtgBr.CellFontBold = True: dtgBr.Text = "名称"
dtgBr.Col = 3: dtgBr.CellFontBold = True: dtgBr.Text = "英文名称"
dtgBr.Col = 4: dtgBr.CellFontBold = True: dtgBr.Text = "原厂编号"
dtgBr.Col = 5: dtgBr.CellFontBold = True: dtgBr.Text = "规格"
dtgBr.Col = 6: dtgBr.CellFontBold = True: dtgBr.Text = "性能参数"
dtgBr.Col = 7: dtgBr.CellFontBold = True: dtgBr.Text = "使用方法"
dtgBr.Col = 8: dtgBr.CellFontBold = True: dtgBr.Text = "适用机组"
dtgBr.Col = 9: dtgBr.CellFontBold = True: dtgBr.Text = "备注"
ii = 0
On Error Resume Next
For oo = 1 To 10000
    dtgN.Row = oo
    dtgN.Col = 1
    If dtgN.Text = "" Then Exit For
    If InStr(dtgN.Text, txtZ.Text) > 0 Then
        ii = ii + 1
        dtgBr.Row = ii
        For kk = 0 To 10
            dtgBr.Col = kk: dtgN.Col = kk
            dtgBr.Text = dtgN.Text
        Next
            GoTo frmGxbjLMxc
    End If
    dtgN.Col = 2
    If InStr(dtgN.Text, txtZ.Text) > 0 Then
        ii = ii + 1
        dtgBr.Row = ii
        For kk = 0 To 10
            dtgBr.Col = kk: dtgN.Col = kk
            dtgBr.Text = dtgN.Text
        Next
                GoTo frmGxbjLMxc
    End If
    dtgN.Col = 4
    If InStr(dtgN.Text, txtZ.Text) > 0 Then
        ii = ii + 1
        dtgBr.Row = ii
        For kk = 0 To 10
            dtgBr.Col = kk: dtgN.Col = kk
            dtgBr.Text = dtgN.Text
        Next
                    GoTo frmGxbjLMxc
    End If
    dtgN.Col = 5
    If InStr(dtgN.Text, txtZ.Text) > 0 Then
        ii = ii + 1
        dtgBr.Row = ii
        For kk = 0 To 10
            dtgBr.Col = kk: dtgN.Col = kk
            dtgBr.Text = dtgN.Text
        Next
                    GoTo frmGxbjLMxc
    End If
    dtgN.Col = 6
    If InStr(dtgN.Text, txtZ.Text) > 0 Then
        ii = ii + 1
        dtgBr.Row = ii
        For kk = 0 To 10
            dtgBr.Col = kk: dtgN.Col = kk
            dtgBr.Text = dtgN.Text
        Next
                    GoTo frmGxbjLMxc
    End If
    dtgN.Col = 7
    If InStr(dtgN.Text, txtZ.Text) > 0 Then
        ii = ii + 1
        dtgBr.Row = ii
        For kk = 0 To 10
            dtgBr.Col = kk: dtgN.Col = kk
            dtgBr.Text = dtgN.Text
        Next
                    GoTo frmGxbjLMxc
    End If
    dtgN.Col = 8
    If InStr(dtgN.Text, txtZ.Text) > 0 Then
        ii = ii + 1
        dtgBr.Row = ii
        For kk = 0 To 10
            dtgBr.Col = kk: dtgN.Col = kk
            dtgBr.Text = dtgN.Text
        Next
                    GoTo frmGxbjLMxc
    End If
    dtgN.Col = 9
    If InStr(dtgN.Text, txtZ.Text) > 0 Then
        ii = ii + 1
        dtgBr.Row = ii
        For kk = 0 To 10
            dtgBr.Col = kk: dtgN.Col = kk
            dtgBr.Text = dtgN.Text
        Next
                GoTo frmGxbjLMxc
    End If
    
frmGxbjLMxc:
    
Next



dtgBr.Visible = True
cmdAgain.Enabled = False
End Sub


Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub cmdC_Click()

Me.tt = "select pb,bh,partname,engName,oName,gg,xn,ff,pb+' '+jz,bz,pid from nlpcool where pb like '%" & txtZ.Text & "%'" & _
        " or bh='" & txtZ.Text & "' or partname like '%" & txtZ.Text & "%' or engname like '%" & txtZ.Text & "%' or oName like '%" & txtZ.Text & "%'" & _
        " or gg like '%" & txtZ.Text & "%' or xn like '%" & txtZ.Text & "%' order by pid"
Call Me.Bound(Me.tt)
On Error Resume Next
txtZ.SelStart = 0
txtZ.SelLength = Len(txtZ.Text)
txtZ.SetFocus
'''''txtZ.SelText = txtZ.Text
cmdAgain.Enabled = True
End Sub

Private Sub dtgBr_DblClick()
On Error Resume Next

Dim tt As String

Dim Ra
Dim ii As Integer
Dim Pid As Long
dtgN.Row = dtgBr.Row
dtgN.Col = 10
Pid = Val(dtgN.Text)
If Pid = 0 Then Exit Sub
On Error GoTo XjaErr
ii = InputBox("请输入数量!")

                                   '添加
    timZm = 1
    On Error Resume Next
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "新询价单"
    mod1.cmd.Parameters("@NBLX") = "市场营销部配件添加"
    mod1.cmd.Parameters("@bh") = frmGxbjNew.lblBh.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    '''''询价单性质
    mod1.cmd.Parameters("@mt2") = "配件"

    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = ii
    mod1.cmd.Parameters("@mm2") = Pid
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
    Else '提交成功,等待系统中心处理数据
        cmdAdd.Enabled = False
        cmdJG.Enabled = False
        Me.Enabled = False
        frmWaitA.Visible = True
        frmWaitA.Timer2.Enabled = False

        frmWaitA.ZOrder 0
        frmWaitA.Timer2.Enabled = True
        timWait.Enabled = True
    End If


    Set mod1.cmd = Nothing
frmStep.Visible = False
frmA.Enabled = True
Exit Sub
XjaErr:

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Call dtgFF
End Sub '

Public Sub dtgFF()
dtgBr.Clear: dtgN.Clear
dtgBr.Rows = 30
dtgBr.Cols = 11
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.CellFontBold = True: dtgBr.Text = "品牌"
dtgBr.Col = 1: dtgBr.CellFontBold = True: dtgBr.Text = "编号"
dtgBr.Col = 2: dtgBr.CellFontBold = True: dtgBr.Text = "名称"
dtgBr.Col = 3: dtgBr.CellFontBold = True: dtgBr.Text = "英文名称"
dtgBr.Col = 4: dtgBr.CellFontBold = True: dtgBr.Text = "原厂编号"
dtgBr.Col = 5: dtgBr.CellFontBold = True: dtgBr.Text = "规格"
dtgBr.Col = 6: dtgBr.CellFontBold = True: dtgBr.Text = "性能参数"
dtgBr.Col = 7: dtgBr.CellFontBold = True: dtgBr.Text = "使用方法"
dtgBr.Col = 8: dtgBr.CellFontBold = True: dtgBr.Text = "适用机组"
dtgBr.Col = 9: dtgBr.CellFontBold = True: dtgBr.Text = "备注"
dtgBr.ColWidth(0) = 0
dtgBr.ColWidth(1) = 870
dtgBr.ColWidth(2) = 1530
dtgBr.ColWidth(3) = 2055
dtgBr.ColWidth(4) = 1410
dtgBr.ColWidth(5) = 2100
dtgBr.ColWidth(6) = -1
dtgBr.ColWidth(7) = 1515
dtgBr.ColWidth(8) = 1500
dtgBr.ColWidth(9) = 1955
dtgBr.ColWidth(10) = 0

dtgN.Rows = 30
dtgN.Cols = 14

End Sub

Public Sub Bound(tt As String)
Dim Ra
Dim La
Dim oo As Integer
Call Me.dtgFF
dtgBr.Visible = False
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly
If mod1.HTP.BOF = True Then
    lblTX.Visible = True
    lblTX.Caption = mod1.chenHu & "," & "很抱歉!货品库中找不到您要的货品,请联系配送中心的,联系电话:18918156727"
Else
    lblTX.Visible = True
    lblTX.Caption = "双击列表中的货品记录,可以向询价单添加该货品!"
End If
On Error Resume Next
Ra = mod1.HTP.GetRows
La = UBound(Ra, 2) + 1
dtgBr.Rows = La + 50
For oo = 0 To La
    dtgBr.Row = oo: dtgN.Row = oo
    dtgBr.Col = 0: dtgBr.Text = Ra(0, oo - 1)
    dtgBr.Col = 1: dtgBr.Text = Ra(1, oo - 1)
    dtgBr.Col = 2: dtgBr.Text = Ra(2, oo - 1)
    dtgBr.Col = 3: dtgBr.Text = Ra(3, oo - 1)
    dtgBr.Col = 4: dtgBr.Text = Ra(4, oo - 1)
    dtgBr.Col = 5: dtgBr.Text = Ra(5, oo - 1)
    dtgBr.Col = 6: dtgBr.Text = Ra(6, oo - 1)
    dtgBr.Col = 7: dtgBr.Text = Ra(7, oo - 1)
    dtgBr.Col = 8: dtgBr.Text = Ra(8, oo - 1)
    dtgBr.Col = 9: dtgBr.Text = Ra(9, oo - 1)
    dtgBr.Col = 10: dtgBr.Text = Ra(10, oo - 1)
'''    dtgBr.Col = 11: dtgBr.Text = Ra(11, oo - 1)
'''    dtgBr.Col = 12: dtgBr.Text = Ra(12, oo - 1)
'''    dtgBr.Col = 13: dtgBr.Text = Ra(13, oo - 1)
    dtgN.Col = 0: dtgN.Text = Ra(0, oo - 1)
    dtgN.Col = 1: dtgN.Text = Ra(1, oo - 1)
    dtgN.Col = 2: dtgN.Text = Ra(2, oo - 1)
    dtgN.Col = 3: dtgN.Text = Ra(3, oo - 1)
    dtgN.Col = 4: dtgN.Text = Ra(4, oo - 1)
    dtgN.Col = 5: dtgN.Text = Ra(5, oo - 1)
    dtgN.Col = 6: dtgN.Text = Ra(6, oo - 1)
    dtgN.Col = 7: dtgN.Text = Ra(7, oo - 1)
    dtgN.Col = 8: dtgN.Text = Ra(8, oo - 1)
    dtgN.Col = 9: dtgN.Text = Ra(9, oo - 1)
    dtgN.Col = 10: dtgN.Text = Ra(10, oo - 1)
'''    dtgN.Col = 11: dtgN.Text = Ra(11, oo - 1)
'''    dtgN.Col = 12: dtgN.Text = Ra(12, oo - 1)
'''    dtgN.Col = 13: dtgN.Text = Ra(13, oo - 1)
Next
dtgBr.Visible = True
End Sub

Private Sub timQuit_Timer()
On Error Resume Next
Dim oo As Integer
Dim jj As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '配件添加
    Call frmGxbjNew.BoundForm(Val(frmGxbjNew.lblBh.ToolTipText))
    Me.Visible = False
End If
timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mm2,mt1,mt2,mt3 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 2 Or timZm = 3 Then
        txtHg.Text = mod1.WP.Fields("mm1").Value
    ElseIf timZm = 6 Then '签名
                lblLc.Caption = mod1.WP.Fields("mm1").Value
                lblFwid.Caption = mod1.WP.Fields("mm2").Value
                lblLcRen.Caption = mod1.WP.Fields("mt1").Value
                lblLcUid.Caption = mod1.WP.Fields("mt2").Value
                lblTX.Caption = "下一流程,将跳至" & mod1.WP.Fields("mt3").Value & ": " & lblLcRen.Caption
                frmQm.Visible = False
    End If
    timWait.Enabled = False
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdJG.Enabled = False
    End If
    timWait.Enabled = False
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdJG.Enabled = False
    End If
    Exit Sub

End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub

Private Sub txtZ_DblClick()
Call cmdAgain_Click
End Sub

Private Sub txtZ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmdC_Click
End If
End Sub
