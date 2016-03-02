VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmLPNew 
   BackColor       =   &H00C0FFC0&
   Caption         =   "新零件库"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6840
      Top             =   30
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6030
      Top             =   30
   End
   Begin VB.Frame frmGy 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1125
      Left            =   180
      TabIndex        =   29
      Top             =   7980
      Width           =   10905
      Begin VB.TextBox txtDj3 
         Height          =   270
         Left            =   8100
         TabIndex        =   41
         Top             =   420
         Width           =   2565
      End
      Begin VB.TextBox txtDj2 
         Height          =   270
         Left            =   4470
         TabIndex        =   40
         Top             =   450
         Width           =   2565
      End
      Begin VB.TextBox txtDj1 
         Height          =   270
         Left            =   840
         TabIndex        =   39
         Top             =   450
         Width           =   2565
      End
      Begin VB.TextBox txtGy3 
         Height          =   270
         Left            =   8100
         TabIndex        =   35
         Top             =   30
         Width           =   2565
      End
      Begin VB.TextBox txtGy2 
         Height          =   270
         Left            =   4470
         TabIndex        =   34
         Top             =   30
         Width           =   2565
      End
      Begin VB.TextBox txtGy1 
         Height          =   270
         Left            =   840
         TabIndex        =   33
         Top             =   30
         Width           =   2565
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "成本价"
         Height          =   225
         Index           =   2
         Left            =   7320
         TabIndex        =   38
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "成本价"
         Height          =   225
         Index           =   1
         Left            =   3645
         TabIndex        =   37
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "成本价"
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   36
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "供应商3"
         Height          =   315
         Left            =   7320
         TabIndex        =   32
         Top             =   60
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "供应商2"
         Height          =   285
         Left            =   3645
         TabIndex        =   31
         Top             =   60
         Width           =   825
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "供应商1"
         Height          =   255
         Left            =   0
         TabIndex        =   30
         Top             =   60
         Width           =   795
      End
   End
   Begin VB.TextBox txtEngName 
      Height          =   315
      Left            =   990
      TabIndex        =   26
      Top             =   6930
      Width           =   2565
   End
   Begin VB.TextBox txtFF 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6300
      TabIndex        =   25
      Top             =   6930
      Width           =   3075
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   285
      Left            =   11970
      TabIndex        =   24
      Top             =   240
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   503
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtZ 
      Height          =   315
      HideSelection   =   0   'False
      Left            =   60
      TabIndex        =   23
      Top             =   90
      Width           =   3885
   End
   Begin VB.CommandButton cmdC 
      BackColor       =   &H00C0FFC0&
      Caption         =   "搜索"
      Height          =   315
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   90
      Width           =   1035
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFC0&
      Caption         =   "返回"
      Height          =   615
      Left            =   13530
      Picture         =   "frmLPNew.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8070
      Width           =   675
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   765
      Left            =   11520
      TabIndex        =   17
      Top             =   8040
      Width           =   2925
      Begin VB.CommandButton cmdCreate 
         BackColor       =   &H00FFFFC0&
         Caption         =   "新建"
         Height          =   645
         Left            =   0
         Picture         =   "frmLPNew.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         Width           =   645
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0FFC0&
         Caption         =   "保存"
         Height          =   645
         Left            =   690
         Picture         =   "frmLPNew.frx":0544
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         Width           =   645
      End
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "删除"
         Enabled         =   0   'False
         Height          =   645
         Left            =   1350
         Picture         =   "frmLPNew.frx":0BAE
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.TextBox txtBz 
      Height          =   645
      Left            =   11190
      TabIndex        =   16
      Top             =   6930
      Width           =   3135
   End
   Begin VB.TextBox txtJz 
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   11190
      TabIndex        =   15
      Top             =   6330
      Width           =   3135
   End
   Begin VB.TextBox txtXn 
      Height          =   315
      Left            =   6300
      TabIndex        =   14
      Top             =   6330
      Width           =   3075
   End
   Begin VB.TextBox txtGG 
      Height          =   315
      Left            =   6300
      TabIndex        =   13
      Top             =   5745
      Width           =   3075
   End
   Begin VB.TextBox txtOName 
      Height          =   315
      Left            =   990
      TabIndex        =   12
      Top             =   7455
      Width           =   2565
   End
   Begin VB.TextBox txtPartName 
      Height          =   315
      Left            =   990
      TabIndex        =   6
      Top             =   6352
      Width           =   2565
   End
   Begin VB.TextBox txtBh 
      Height          =   315
      Left            =   990
      TabIndex        =   5
      Top             =   5745
      Width           =   2565
   End
   Begin VB.TextBox txtPb 
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   11190
      TabIndex        =   2
      Top             =   5745
      Width           =   3105
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   5025
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   8864
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "英文名称"
      Height          =   285
      Left            =   180
      TabIndex        =   28
      Top             =   6960
      Width           =   945
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "使用方法"
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4740
      TabIndex        =   27
      Top             =   6960
      Width           =   1245
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "备注"
      Height          =   315
      Left            =   10020
      TabIndex        =   11
      Top             =   6960
      Width           =   585
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "适用机组"
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   10020
      TabIndex        =   10
      Top             =   6390
      Width           =   1185
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "性能参数"
      Height          =   315
      Left            =   4740
      TabIndex        =   9
      Top             =   6367
      Width           =   1245
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "规格"
      Height          =   315
      Left            =   4740
      TabIndex        =   8
      Top             =   5775
      Width           =   1305
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "原厂编号"
      Height          =   315
      Left            =   180
      TabIndex        =   7
      Top             =   7500
      Width           =   795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "名称"
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   6367
      Width           =   825
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "编号"
      Height          =   285
      Left            =   180
      TabIndex        =   3
      Top             =   5775
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "品牌类"
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   10080
      TabIndex        =   1
      Top             =   5775
      Width           =   615
   End
End
Attribute VB_Name = "frmLPNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tt As String
Dim timZm As Integer

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
End Sub

Public Sub dtgFF()
dtgBr.Clear
dtgN.Clear
dtgBr.Rows = 30
dtgBr.Cols = 18
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
dtgBr.ColWidth(11) = 0
dtgBr.ColWidth(12) = 0
dtgBr.ColWidth(13) = 0
dtgBr.ColWidth(14) = 0
dtgBr.ColWidth(15) = 0
dtgBr.ColWidth(16) = 0
dtgBr.ColWidth(17) = 0
dtgN.Rows = 30
dtgN.Cols = 18

End Sub

Private Sub cmdC_Click()

Me.tt = "select pb,bh,partname,engName,oName,gg,xn,ff,pb+' '+jz,bz,pid,gid1,dj1,gid2,dj2,gid3,dj3,JZ from nlpcool where pb like '%" & txtZ.Text & "%'" & _
        " or bh='" & txtZ.Text & "' or partname like '%" & txtZ.Text & "%' or engname like '%" & txtZ.Text & "%' or oName like '%" & txtZ.Text & "%'" & _
        " or gg like '%" & txtZ.Text & "%' or xn like '%" & txtZ.Text & "%' or jz like '%" & txtZ.Text & "%' order by pid"
Call Me.Bound(Me.tt)
dtgBr.Row = 1
Call dtgBr_Click
End Sub

Private Sub cmdCreate_Click()
Call dtgFF
Call Me.MQing
frmGY.Enabled = False
End Sub


Private Sub cmdDel_Click()
On Error Resume Next
Dim ii As Integer
Dim tt As String
If Val(txtBh.ToolTipText) = 0 Then Exit Sub
ii = MsgBox("是否删除此货品记录?", vbQuestion + vbYesNo + vbDefaultButton2, "请确认!")
If ii = vbNo Then Exit Sub
    timZm = 2 '删除
    On Error Resume Next
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "货品资料"
    mod1.cmd.Parameters("@NBLX") = "删除"
    mod1.cmd.Parameters("@bh") = txtBh.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtBh.Text
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = 0
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
End Sub

Private Sub cmdSave_Click()
On Error Resume Next

Dim tt As String
    timZm = 1
    On Error Resume Next
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "货品资料"
    mod1.cmd.Parameters("@NBLX") = "编辑"
    mod1.cmd.Parameters("@bh") = txtBh.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = txtBh.Text
    mod1.cmd.Parameters("@mt2") = txtPartName.Text
    mod1.cmd.Parameters("@mt3") = txtEngName.Text
    mod1.cmd.Parameters("@mt4") = txtOname.Text
    mod1.cmd.Parameters("@mt5") = txtGG.Text
    mod1.cmd.Parameters("@mt6") = txtXN.Text
    mod1.cmd.Parameters("@mt7") = txtFF.Text
    mod1.cmd.Parameters("@mt8") = txtPb.Text
    mod1.cmd.Parameters("@mt9") = txtJz.Text
    mod1.cmd.Parameters("@mt10") = txtBz.Text
    mod1.cmd.Parameters("@mt11") = txtGy1.Text
    mod1.cmd.Parameters("@mt12") = txtGy2.Text
    mod1.cmd.Parameters("@mt13") = txtGY3.Text
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txtDj1.Text)
    mod1.cmd.Parameters("@mm2") = Val(txtDj2.Text)
    mod1.cmd.Parameters("@mm3") = Val(txtDj3.Text)
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


End Sub


Private Sub dtgBr_Click()
Dim Pid As Long
dtgN.Row = dtgBr.Row
dtgN.Col = 10
Pid = Val(dtgN.Text)
If Pid = 0 Then Exit Sub
Call Me.MQing
tt = "select pb,bh,partname,engName,oName,gg,xn,ff,pb+' '+jz,bz,pid,gid1,dj1,gid2,dj2,gid3,dj3,JZ from nlpcool where pb like '%" & txtZ.Text & "%'" & _
        " or bh='" & txtZ.Text & "' or partname like '%" & txtZ.Text & "%' or engname like '%" & txtZ.Text & "%' or oName like '%" & txtZ.Text & "%'" & _
        " or gg like '%" & txtZ.Text & "%' or xn like '%" & txtZ.Text & "%' order by pid"
dtgN.Col = 0: txtPb.Text = dtgN.Text
dtgN.Col = 1: txtBh.Text = dtgN.Text
dtgN.Col = 10: txtBh.ToolTipText = dtgN.Text
dtgN.Col = 2: txtPartName.Text = dtgN.Text
dtgN.Col = 3: txtEngName.Text = dtgN.Text
dtgN.Col = 4: txtOname.Text = dtgN.Text
dtgN.Col = 5: txtGG.Text = dtgN.Text
dtgN.Col = 6: txtXN.Text = dtgN.Text
dtgN.Col = 7: txtFF.Text = dtgN.Text
dtgN.Col = 17: txtJz.Text = dtgN.Text
dtgN.Col = 9: txtBz.Text = dtgN.Text
dtgN.Col = 11: txtGy1.Text = dtgN.Text
dtgN.Col = 13: txtGy2.Text = dtgN.Text
dtgN.Col = 15: txtGY3.Text = dtgN.Text
dtgN.Col = 12: txtDj1.Text = dtgN.Text
dtgN.Col = 14: txtDj2.Text = dtgN.Text
dtgN.Col = 16: txtDj3.Text = dtgN.Text
End Sub


Private Sub Form_Load()
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
If mod1.DName = "马晓聪" Or mod1.DName = "宋晓炯" Or mod1.DName = "" Or Ywy = "吴金荣" Or mod1.DName = "杨晓刚" Then
    frmGY.Visible = True
Else
    frmGY.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Visible = False
frmZu.Enabled = True
End Sub

Public Sub Bound(tt As String)
Dim Ra
Dim La
Dim oo As Integer
Call Me.dtgFF
dtgBr.Visible = False
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly

On Error Resume Next
Ra = mod1.HTP.GetRows
La = UBound(Ra, 2) + 1
dtgBr.Rows = La + 50
dtgN.Rows = dtgBr.Rows
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
    dtgBr.Col = 11: dtgBr.Text = Ra(11, oo - 1)
    dtgBr.Col = 12: dtgBr.Text = Ra(12, oo - 1)
    dtgBr.Col = 13: dtgBr.Text = Ra(13, oo - 1)
    dtgBr.Col = 14: dtgBr.Text = Ra(14, oo - 1)
    dtgBr.Col = 15: dtgBr.Text = Ra(15, oo - 1)
    dtgBr.Col = 16: dtgBr.Text = Ra(16, oo - 1)
    dtgBr.Col = 17: dtgBr.Text = Ra(17, oo - 1)
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
    dtgN.Col = 11: dtgN.Text = Ra(11, oo - 1)
    dtgN.Col = 12: dtgN.Text = Ra(12, oo - 1)
    dtgN.Col = 13: dtgN.Text = Ra(13, oo - 1)
    dtgN.Col = 14: dtgN.Text = Ra(14, oo - 1)
    dtgN.Col = 15: dtgN.Text = Ra(15, oo - 1)
    dtgN.Col = 16: dtgN.Text = Ra(16, oo - 1)
    dtgN.Col = 17: dtgN.Text = Ra(17, oo - 1)
Next
dtgBr.Visible = True
End Sub

Public Sub MQing()
txtPb.Text = ""
txtBh.Text = ""
txtBh.ToolTipText = ""
txtPartName.Text = ""
txtEngName.Text = ""
txtOname.Text = ""
txtGG.Text = ""
txtXN.Text = ""
txtFF.Text = ""
txtJz.Text = ""
txtBz.Text = ""
txtGy1.Text = ""
txtGy2.Text = ""
txtGY3.Text = ""
txtDj1.Text = ""
txtDj2.Text = ""
txtDj3.Text = ""
End Sub

Private Sub timQuit_Timer()
On Error Resume Next
Dim oo As Integer
Dim jj As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 1 Then '编辑
    dtgN.Row = 1: dtgN.Col = 1 '如果为添加,则列表显示新记录
    If dtgN.Text = "" Then
        Me.tt = "select pb,bh,partname,engName,oName,gg,xn,ff,pb+' '+jz,bz,pid,gid1,dj1,gid2,dj2,gid3,dj3,JZ from nlpcool where pid=" & Val(txtBh.ToolTipText)
    End If
    Call Me.Bound(Me.tt)
    frmGY.Enabled = True
ElseIf timZm = 2 Then '删除
    Call Me.dtgFF
    Call Me.MQing
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
    If timZm = 1 Then
        txtBh.ToolTipText = mod1.WP.Fields("mm1").Value
'''''''    ElseIf timZm = 6 Then '签名
'''''''                lblLc.Caption = mod1.WP.Fields("mm1").Value
'''''''                lblFwid.Caption = mod1.WP.Fields("mm2").Value
'''''''                lblLcRen.Caption = mod1.WP.Fields("mt1").Value
'''''''                lblLcUid.Caption = mod1.WP.Fields("mt2").Value
'''''''                lblTx.Caption = "下一流程,将跳至" & mod1.WP.Fields("mt3").Value & ": " & lblLcRen.Caption
'''''''                frmQm.Visible = False
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

Private Sub txtZ_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Call cmdC_Click
End If
End Sub


