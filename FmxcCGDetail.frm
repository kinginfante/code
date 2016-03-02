VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcCGDetail 
   BackColor       =   &H00C0FFC0&
   Caption         =   "新版采购单"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgGy 
      Height          =   1335
      Left            =   1410
      TabIndex        =   30
      Top             =   1080
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   2355
      _Version        =   393216
      BackColor       =   12648384
      Rows            =   50
      FixedCols       =   0
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
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   90
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印"
      Height          =   645
      Left            =   12510
      Picture         =   "FmxcCGDetail.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8400
      Width           =   675
   End
   Begin VB.TextBox txtBh 
      Height          =   270
      Left            =   11490
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   1110
      Width           =   3315
   End
   Begin VB.TextBox txtGy 
      Height          =   270
      Left            =   1410
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   780
      Width           =   3315
   End
   Begin VB.TextBox txtQy 
      Height          =   270
      Left            =   6660
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   750
      Width           =   3315
   End
   Begin VB.TextBox txtDRQ 
      Height          =   270
      Left            =   11490
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text3"
      Top             =   750
      Width           =   3315
   End
   Begin VB.TextBox txtADR 
      Height          =   270
      Left            =   6660
      TabIndex        =   16
      Text            =   "Text4"
      Top             =   1110
      Width           =   3315
   End
   Begin VB.TextBox txtLLR 
      Height          =   270
      Left            =   1410
      TabIndex        =   15
      Text            =   "Text5"
      Top             =   1125
      Width           =   3315
   End
   Begin VB.CommandButton cmdback 
      BackColor       =   &H00FFFFC0&
      Caption         =   "返回"
      Height          =   645
      Left            =   14520
      Picture         =   "FmxcCGDetail.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8400
      Width           =   645
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00C0FFC0&
      Caption         =   "保存"
      Height          =   645
      Left            =   13860
      Picture         =   "FmxcCGDetail.frx":076C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8400
      Width           =   645
   End
   Begin VB.CommandButton cmdCreate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "新建"
      Height          =   645
      Left            =   13080
      Picture         =   "FmxcCGDetail.frx":0DD6
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7110
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H00FFFFC0&
      Caption         =   "修改"
      Height          =   645
      Left            =   13200
      Picture         =   "FmxcCGDetail.frx":1218
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "修改"
      Top             =   8400
      Width           =   645
   End
   Begin VB.CommandButton cmdNQ 
      BackColor       =   &H008080FF&
      Caption         =   "审核"
      Height          =   645
      Left            =   11820
      Picture         =   "FmxcCGDetail.frx":1522
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8400
      Width           =   675
   End
   Begin VB.Frame frmQm 
      BackColor       =   &H00C0FFC0&
      Caption         =   "评审建议"
      ForeColor       =   &H000000FF&
      Height          =   1785
      Left            =   60
      TabIndex        =   2
      Top             =   7170
      Visible         =   0   'False
      Width           =   6315
      Begin VB.CommandButton cmdDing 
         BackColor       =   &H00FF8080&
         Caption         =   "决定"
         Height          =   285
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton optT2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "拒绝"
         Height          =   195
         Left            =   5220
         TabIndex        =   5
         Top             =   870
         Width           =   675
      End
      Begin VB.OptionButton OptT1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "同意"
         Height          =   225
         Left            =   5220
         TabIndex        =   4
         Top             =   480
         Width           =   705
      End
      Begin VB.TextBox txtQM 
         BackColor       =   &H00C0FFFF&
         Height          =   1365
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   300
         Width           =   4965
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgP 
      Height          =   2865
      Left            =   30
      TabIndex        =   7
      Top             =   6240
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   5054
      _Version        =   393216
      BackColor       =   15728356
      ForeColor       =   8404992
      Rows            =   15
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   16777152
      ForeColorFixed  =   0
      BackColorBkg    =   15728356
      GridColorFixed  =   8404992
      GridColorUnpopulated=   8404992
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   4335
      Left            =   30
      TabIndex        =   14
      Top             =   1530
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   7646
      _Version        =   393216
      BackColor       =   16777152
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   16777152
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtPRQ 
      Height          =   270
      Left            =   11490
      TabIndex        =   20
      Top             =   750
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   476
      _Version        =   393216
      CalendarBackColor=   8454016
      CalendarTitleBackColor=   16711808
      CalendarTrailingForeColor=   -2147483635
      Format          =   114491393
      CurrentDate     =   38797
   End
   Begin VB.OLE OLE1 
      Height          =   645
      Left            =   7860
      TabIndex        =   29
      Top             =   6540
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "编号"
      Height          =   285
      Left            =   10830
      TabIndex        =   26
      Top             =   1140
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "供应商"
      Height          =   255
      Left            =   390
      TabIndex        =   25
      Top             =   810
      Width           =   885
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "联系人"
      Height          =   255
      Left            =   390
      TabIndex        =   24
      Top             =   1185
      Width           =   885
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "签订地点"
      Height          =   255
      Left            =   5490
      TabIndex        =   23
      Top             =   780
      Width           =   885
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "签订日期"
      Height          =   255
      Left            =   10500
      TabIndex        =   22
      Top             =   780
      Width           =   885
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "收货地址"
      Height          =   255
      Left            =   5490
      TabIndex        =   21
      Top             =   1110
      Width           =   885
   End
   Begin VB.Label lblTx 
      BackStyle       =   0  'Transparent
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   7530
      TabIndex        =   8
      Top             =   8610
      Width           =   3585
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "采  购  合  同"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   405
      Left            =   5760
      TabIndex        =   1
      Top             =   270
      Width           =   2865
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "上海杰升商贸有限公司"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   465
      Left            =   2190
      TabIndex        =   0
      Top             =   270
      Width           =   3735
   End
End
Attribute VB_Name = "FmxcCGDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lc As Integer
Dim LCRen As String
Dim LCUid As String
Dim Fwid As Long
Dim LL As String
Dim LLUid As String

Public Sub dtgGYFF()
dtgGy.Clear
dtgGy.Rows = 50
dtgGy.Cols = 2
dtgGy.Row = 0
dtgGy.Col = 0: dtgGy.Text = "供应商名称（鼠标双击选择）": dtgGy.CellFontBold = True
dtgGy.ColWidth(1) = 0
dtgGy.ColWidth(0) = 3480

End Sub
Public Sub Qing()
txtGy.Text = ""
txtGy.ToolTipText = ""
txtQy.Text = ""
txtLLR.Text = ""
txtADR.Text = ""
dtPRQ.Visible = False
txtBh.Text = ""
txtDRQ.Text = ""
lblTx.Caption = ""
dtgGy.Visible = False
Call Me.dtgBRFF
Call Me.dtgPFF
End Sub

Public Sub dtgPFF()
Dim oo As Integer
For oo = 1 To dtgP.Rows - 1
    dtgP.RowHeight(oo) = dtgP.RowHeight(0)
Next
dtgP.Clear
dtgP.Row = 0
dtgP.Col = 0: dtgP.Text = "日期": dtgP.Col = 1: dtgP.Text = "姓名": dtgP.Col = 2: dtgP.Text = "职能": dtgP.Col = 3: dtgP.Text = "评审建议": dtgP.Col = 4: dtgP.Text = "审核":
dtgP.ColWidth(0) = 1005
dtgP.ColWidth(1) = 1005
dtgP.ColWidth(2) = 0
 dtgP.ColWidth(3) = 3000: dtgP.ColWidth(4) = 1035
For oo = 0 To 4
    dtgP.Col = oo
    dtgP.CellFontBold = True
Next
End Sub

Public Sub dtgBRFF()
dtgBr.Clear
dtgBr.Cols = 8
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "货品编号": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "货品名称": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "描述": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "单价": dtgBr.CellFontBold = True
dtgBr.Col = 4: dtgBr.Text = "数量": dtgBr.CellFontBold = True
dtgBr.Col = 5: dtgBr.Text = "小计": dtgBr.CellFontBold = True
dtgBr.Col = 6: dtgBr.Text = "对应合同": dtgBr.CellFontBold = True
dtgBr.ColWidth(1) = 2040
dtgBr.ColWidth(2) = 7695
dtgBr.ColWidth(7) = 0
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub cmdMod_Click()
cmdSave.Enabled = True
dtPRQ.Visible = True
End Sub

Private Sub cmdPrint_Click()

'''Dim bt() As Byte
'''Dim tt As String
'''On Error Resume Next
'''Kill "c:\work\*.xls": Kill "c:\work\*.doc"
'''tt = "select fnr,fsize,fname from ht where fid=" & Val(txtHtbh.ToolTipText)
'''frmGGL.adoFile.Recordset.Close
'''frmGGL.adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockReadOnly, adCmdText
'''ReDim bt(frmGGL.adoFile.Recordset.Fields("Fsize").Value) As Byte
'''bt() = frmGGL.adoFile.Recordset.Fields("FNR").GetChunk(frmGGL.adoFile.Recordset.Fields("Fsize").Value + 1)
'''
'''Open ("c:\work\" & frmGGL.adoFile.Recordset.Fields("fname").Value) For Binary As #2
'''Put #2, , bt()
'''Close #2
'''
'''    frmGGL.OLE2.SourceDoc = "c:\work\" & frmGGL.adoFile.Recordset.Fields("fname").Value
'''    frmGGL.OLE2.Action = 1
'''    frmGGL.OLE2.DoVerb (-2)
Dim tt As String
Dim Ra
tt = "update cou set cou=" & Val(txtBh.Text) & ";select @@identity"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workBD, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = mod1.HTP.NextRecordset
Ra = mod1.HTP.GetRows
Set mod1.HTP = Nothing
On Error Resume Next
    OLE1.SourceDoc = "c:\work\合同样本.xls"
    OLE1.Action = 1
    OLE1.DoVerb (-2)
End Sub

Private Sub cmdSave_Click()
Dim tt As String
Dim ii As Integer
Dim Ra
Dim Rb

    timZm = 1 '保存
        
        Set mod1.HTP = Nothing
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.CC
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "采购合同"
        mod1.cmd.Parameters("@NBLX") = "保存"
        mod1.cmd.Parameters("@bh") = Trim(txtBh.Text)
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = txtLLR.Text
        mod1.cmd.Parameters("@mt2") = txtQy.Text
        mod1.cmd.Parameters("@mt3") = txtADR.Text
        mod1.cmd.Parameters("@mt11") = ""
        mod1.cmd.Parameters("@mm1") = Val(txtGy.ToolTipText)
        mod1.cmd.Parameters("@mb1") = Null
        If txtDRQ.Text <> "" Then
             mod1.cmd.Parameters("@md1") = txtDRQ.Text
        Else
             mod1.cmd.Parameters("@md1") = Null
        End If
        Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
        mod1.cmd.Execute
        mod1.Zid = mod1.cmd.Parameters("@zid").Value
        If mod1.cmd.Parameters("@errch").Value <> "成功" Then
            MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
                cmdSave.Enabled = False
            Exit Sub
        Else '提交成功,等待系统中心处理数据
            Me.Enabled = False
            frmWaitA.Visible = True
            frmWaitA.Timer2.Enabled = False
    
            frmWaitA.ZOrder 0
            frmWaitA.Timer2.Enabled = True
            timWait.Enabled = True
            cmdSave.Enabled = False
        End If
    Set mod1.cmd = Nothing

End Sub

Private Sub dtgGy_Click()
On Error Resume Next


    dtgGy.Col = 0: txtGy.Text = dtgGy.Text
    dtgGy.Col = 1: txtGy.ToolTipText = dtgGy.Text

End Sub

Private Sub dtPRQ_CloseUp()
Me.txtDRQ.Text = dtPRQ.Value
End Sub


Private Sub Form_Click()
dtgGy.Visible = False
txtGy.Locked = False
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
dtPRQ.Value = Date
End Sub


Public Sub Bound(Cid As Long)
Dim Ra, Rb, RC
Dim Lb As Integer
Dim oo As Integer
Call Qing
tt = "declare @gyid int;" & _
    "select @gyid=gyid from cgd where cid=" & Cid & ";" & _
    "select cid,gyid,qy,llr,drq,adr,lc,lcren,lcuid,fwid,ll,lluid from CGD where cid=" & Cid & ";" & _
    "select bh,partname,detail,dj,sl,hg,hid from CGDDT where cid=" & Cid & ";" & _
    "select mc from gymxc where gid=@gyid"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing

Me.txtBh.Text = Ra(0, 0): Me.txtBh.ToolTipText = Cid
Me.txtGy.ToolTipText = Ra(1, 0)
Me.txtGy.Text = RC(0, 0)
Me.txtQy.Text = Ra(2, 0)
Me.txtLLR.Text = Ra(3, 0)
Me.txtDRQ.Text = Ra(4, 0)
Me.txtADR.Text = Ra(5, 0)
Lc = Ra(6, 0)
LCRen = Ra(7, 0)
LCUid = Ra(8, 0)
Fwid = Ra(9, 0)
LL = Ra(10, 0)
LLUid = Ra(11, 0)

Lb = UBound(Rb, 2) + 1
dtgBr.Rows = Lb + 50
For oo = 1 To Lb
    dtgBr.Row = oo
    dtgBr.Col = 0: dtgBr.Text = Rb(0, oo - 1)
    dtgBr.Col = 1: dtgBr.Text = Rb(1, oo - 1)
    dtgBr.Col = 2: dtgBr.Text = Rb(2, oo - 1)
    dtgBr.Col = 3: dtgBr.Text = Rb(3, oo - 1)
    dtgBr.Col = 4: dtgBr.Text = Rb(4, oo - 1)
    dtgBr.Col = 5: dtgBr.Text = Rb(5, oo - 1)
    dtgBr.Col = 6: dtgBr.Text = Rb(6, oo - 1)
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
If timZm = 1 Then '保存

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

Private Sub txtGy_DblClick()
dtgGy.Visible = True
txtGy.Locked = False
Call Me.dtgGYFF
End Sub


Private Sub txtGy_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tt As String
Dim Ra
Dim La As Long
Dim oo As Long
If Len(txtGy.Text) < 2 Then Exit Sub
'tt = "select mc,gid from gymxc where mc like '%" & txtGy.Text & "%' and delf=1 and lc=100"
tt = "select mc,gid from gymxc where mc like '%" & txtGy.Text & "%' and delf=1"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Call Me.dtgGYFF
For oo = 1 To La
    dtgGy.Row = oo
    dtgGy.Col = 0: dtgGy.Text = Ra(0, oo - 1)
    dtgGy.Col = 1: dtgGy.Text = Ra(1, oo - 1)
Next
End Sub


