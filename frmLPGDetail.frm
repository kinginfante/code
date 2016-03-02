VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmLPGDetail 
   BackColor       =   &H00FFFFC0&
   Caption         =   "零配件资料"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13995
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   13995
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   810
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   765
      Left            =   240
      TabIndex        =   33
      Top             =   2850
      Width           =   3855
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "删除"
         Enabled         =   0   'False
         Height          =   645
         Left            =   1350
         Picture         =   "frmLPGDetail.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   0
         Width           =   645
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0FFC0&
         Caption         =   "保存"
         Height          =   645
         Left            =   690
         Picture         =   "frmLPGDetail.frx":018A
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   0
         Width           =   645
      End
      Begin VB.CommandButton cmdCreate 
         BackColor       =   &H00FFFFC0&
         Caption         =   "新建"
         Height          =   645
         Left            =   0
         Picture         =   "frmLPGDetail.frx":07F4
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.Frame frmRealNumbers 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3495
      Left            =   6180
      TabIndex        =   25
      Top             =   2820
      Width           =   6495
      Begin VB.CommandButton cmd1 
         Caption         =   "HM"
         Height          =   585
         Left            =   4770
         Picture         =   "frmLPGDetail.frx":0C36
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2880
         Width           =   675
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgRealNumbers 
         Height          =   2355
         Left            =   0
         TabIndex        =   26
         Top             =   390
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   4154
         _Version        =   393216
         BackColor       =   16777152
         Rows            =   10
         FixedCols       =   0
         BackColorFixed  =   15728356
         BackColorBkg    =   16777152
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblNumbers 
         BackStyle       =   0  'Transparent
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         TabIndex        =   32
         Top             =   60
         Width           =   4095
      End
      Begin VB.Label lblB1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label10"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         TabIndex        =   31
         Top             =   90
         Width           =   1785
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   5490
      Picture         =   "frmLPGDetail.frx":0D20
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Width           =   675
   End
   Begin VB.Frame frmSupplier 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3495
      Left            =   6810
      TabIndex        =   28
      Top             =   2070
      Width           =   6495
      Begin VB.CommandButton cmd2 
         Caption         =   "HM"
         Height          =   585
         Left            =   4260
         Picture         =   "frmLPGDetail.frx":0E22
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2910
         Width           =   675
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgSupplier 
         Height          =   2625
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   4630
         _Version        =   393216
         BackColor       =   16777152
         Rows            =   10
         FixedCols       =   0
         BackColorFixed  =   15728356
         BackColorBkg    =   16777152
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.TextBox txtReplaceNumber2 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00C000C0&
      Height          =   285
      Left            =   4470
      TabIndex        =   21
      Top             =   2340
      Width           =   1635
   End
   Begin VB.TextBox txtReplaceNumber1 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1470
      TabIndex        =   20
      Top             =   2325
      Width           =   1635
   End
   Begin VB.TextBox txtOriginallyNumbers 
      BackColor       =   &H00EFFEE4&
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   1470
      TabIndex        =   19
      Top             =   1755
      Width           =   4635
   End
   Begin VB.TextBox txtPinYin 
      BackColor       =   &H00EFFEE4&
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   5010
      TabIndex        =   18
      Top             =   810
      Width           =   1095
   End
   Begin VB.TextBox txtPartName 
      BackColor       =   &H00EFFEE4&
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   1470
      TabIndex        =   16
      Top             =   1230
      Width           =   4635
   End
   Begin VB.ComboBox txtPartsCategory2 
      BackColor       =   &H00EFFEE4&
      ForeColor       =   &H00004000&
      Height          =   300
      Left            =   4290
      TabIndex        =   15
      Text            =   "Combo5"
      Top             =   810
      Width           =   720
   End
   Begin VB.ComboBox txtPartsCategory1 
      BackColor       =   &H00EFFEE4&
      ForeColor       =   &H00004000&
      Height          =   300
      Left            =   3600
      TabIndex        =   14
      Text            =   "Combo4"
      Top             =   810
      Width           =   705
   End
   Begin VB.ComboBox txtUnitModel 
      BackColor       =   &H00EFFEE4&
      ForeColor       =   &H00004000&
      Height          =   300
      Left            =   2400
      TabIndex        =   13
      Text            =   "Combo3"
      Top             =   810
      Width           =   1185
   End
   Begin VB.ComboBox txtUnitSeries 
      BackColor       =   &H00EFFEE4&
      ForeColor       =   &H00004000&
      Height          =   300
      Left            =   1470
      TabIndex        =   12
      Text            =   "Combo2"
      Top             =   810
      Width           =   945
   End
   Begin VB.ComboBox txtUnitBrand 
      BackColor       =   &H00EFFEE4&
      ForeColor       =   &H00004000&
      Height          =   300
      ItemData        =   "frmLPGDetail.frx":0F0C
      Left            =   570
      List            =   "frmLPGDetail.frx":0F1F
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   810
      Width           =   885
   End
   Begin VB.TextBox txtHMNumbers 
      BackColor       =   &H00EFFEE4&
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   8250
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   810
      Width           =   4635
   End
   Begin VB.Label lblPP 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2460
      TabIndex        =   24
      Top             =   90
      Width           =   1455
   End
   Begin VB.Image imgPart 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2565
      Left            =   6600
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   3225
   End
   Begin VB.Label lblRealPart 
      BackStyle       =   0  'Transparent
      Caption         =   "Label14"
      Height          =   255
      Left            =   8010
      TabIndex        =   23
      Top             =   3450
      Width           =   4395
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "实物照片"
      Height          =   285
      Left            =   6720
      TabIndex        =   22
      Top             =   3480
      Width           =   1005
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "拼音编号"
      ForeColor       =   &H00004000&
      Height          =   225
      Left            =   5010
      TabIndex        =   17
      Top             =   540
      Width           =   765
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLPGDetail.frx":0F49
      ForeColor       =   &H00C000C0&
      Height          =   315
      Left            =   3330
      TabIndex        =   8
      Top             =   2370
      Width           =   1125
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmLPGDetail.frx":0F5B
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   -540
      TabIndex        =   7
      Top             =   2370
      Width           =   1875
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "原厂零件编号"
      ForeColor       =   &H00004000&
      Height          =   315
      Left            =   -210
      TabIndex        =   6
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "零件名称"
      ForeColor       =   &H00004000&
      Height          =   315
      Left            =   -60
      TabIndex        =   5
      Top             =   1260
      Width           =   1395
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "豪曼零件编号"
      ForeColor       =   &H00004000&
      Height          =   315
      Left            =   6930
      TabIndex        =   4
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "机组类别"
      ForeColor       =   &H00004000&
      Height          =   225
      Left            =   3780
      TabIndex        =   3
      Top             =   540
      Width           =   915
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "机组型号"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   540
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "机组系列"
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   540
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "机组品牌"
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   570
      TabIndex        =   0
      Top             =   540
      Width           =   885
   End
End
Attribute VB_Name = "frmLPGDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd1_Click()
Me.frmRealNumbers.Visible = False
End Sub

Private Sub cmd2_Click()
frmSupplier.Visible = False
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
End Sub





Private Sub cmdCreate_Click()
Call Me.Initialize
End Sub

Private Sub cmdSave_Click()
Dim tt As String
If Me.txtUnitBrand.Text = "" Or Me.txtUnitSeries.Text = "" Or Me.txtUnitModel.Text = "" Or Me.txtPartsCategory1.Text = "" Or Me.txtPartsCategory2.Text = "" Or Me.txtPartName.Text = "" Or _
     Me.txtOriginallyNumbers.Text = "" Then
    MsgBox "配件资料不全!"
    Exit Sub
End If

On Error Resume Next

timZm = 1 '保存
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "零件事业部配件"
    mod1.cmd.Parameters("@NBLX") = "保存"
    mod1.cmd.Parameters("@bh") = Me.txtHMNumbers.ToolTipText
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = lblPP.Caption
    mod1.cmd.Parameters("@mt2") = Me.txtUnitBrand.Text
    mod1.cmd.Parameters("@mt3") = Me.txtUnitSeries.Text
    mod1.cmd.Parameters("@mt4") = Me.txtUnitModel.Text
    mod1.cmd.Parameters("@mt5") = Me.txtPartsCategory1.Text
    mod1.cmd.Parameters("@mt6") = Me.txtPartsCategory2.Text
    mod1.cmd.Parameters("@mt7") = Me.txtPartName.Text
    mod1.cmd.Parameters("@mt8") = Me.txtPinYin.Text
    mod1.cmd.Parameters("@mt9") = Me.txtHMNumbers.Text
    mod1.cmd.Parameters("@mt10") = Me.txtOriginallyNumbers.Text
    mod1.cmd.Parameters("@mt11") = Me.txtReplaceNumber1.Text
    mod1.cmd.Parameters("@mt12") = Me.txtReplaceNumber2.Text
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = Null
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
        
    End If







   
Set mod1.cmd = Nothing
End Sub

Private Sub dtgRealNumbers_DblClick()
frmSupplier.Visible = True

End Sub


Private Sub dtgSupplier_DblClick()
frmGyDetail.Show
frmGyDetail.ZOrder 0
End Sub


Private Sub Form_Load()
Me.Height = 4005
Me.Width = 6585
Me.Left = 0
Me.Top = 0
Call Me.Initialize
frmRealNumbers.Left = 0
frmRealNumbers.Top = 0
frmRealNumbers.Visible = False
frmSupplier.Top = 0
frmSupplier.Left = 0
frmSupplier.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Cancel = True
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub
Public Function GetPY(strHZ As String) As String '获得单个汉字拼音的首字符
    strHZ = Hex(Asc(strHZ))    '将汉字转换为其内码的十六进制字符串
    Select Case strHZ
    Case "B0A1" To "B0C4"
        GetPY = "a"
    Case "B0C5" To "B2C0"
        GetPY = "b"
    Case "B2C1" To "B4ED"
        GetPY = "c"
    Case "B4EE" To "B6E9"
      GetPY = "d"
    Case "B6EA" To "B7A1"
        GetPY = "e"
    Case "B7A2" To "B8C0"
        GetPY = "f"
    Case "B8C1" To "B9FD"
        GetPY = "g"
    Case "B9FE" To "BBF6"
        GetPY = "h"
    Case "BBF7" To "BFA5"
        GetPY = "j"
    Case "BFA6" To "C0AB"
        GetPY = "k"
    Case "C0AC" To "C2E7"
        GetPY = "l"
    Case "C2E8" To "C4C2"
        GetPY = "m"
    Case "C4C3" To "C5B5"
        GetPY = "n"
    Case "C5B6" To "C5BD"
        GetPY = "o"
    Case "C5BE" To "C6D9"
        GetPY = "p"
    Case "C6DA" To "C8BA"
        GetPY = "q"
    Case "C8BB" To "C8F5"
        GetPY = "r"
    Case "C8F6" To "CBF9"
        GetPY = "s"
    Case "CBFA" To "CDD9"
        GetPY = "t"
    Case "CDDA" To "CEF3"
        GetPY = "w"
    Case "CEF4" To "D188"
        GetPY = "x"
    Case "D189" To "D4D0"
        GetPY = "y"
    Case "D4D1" To "D7F9"
        GetPY = "z"
    Case Else
        GetPY = " "
    End Select
End Function

Public Function GetCode(strZF) As String  '将汉字字符串转换为其拼音的首字符串
    If strZF = "" Then Exit Function
    Dim i As Integer, S As String
    For i = 1 To Len(strZF)
        S = Mid(strZF, i, 1)
        GetCode = GetCode & GetPY(S)
    Next i
End Function

Private Sub timQuit_Timer()
Dim oo As Integer
Dim ii As Integer
On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
Dim tt As String
If timZm = 1 Then '保存
ElseIf timZm = 2 Then '签字
'''    cmdDing.Enabled = True
'''    txtQM.Text = ""
'''    frmQm.Visible = False
'''    lblTX.Visible = True
'''    timQuit.Enabled = False
'''    If Dialog.Visible = True Then
'''        Call mod1.refEnvent(1)
'''    End If
End If
timQuit.Enabled = False
Me.WindowState = 0
Me.ZOrder 0
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mt1,mm2,mt2,mt3 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0
    If timZm = 1 Then
        cmdSave.Enabled = False
        Me.txtHMNumbers.ToolTipText = mod1.WP.Fields("mm1").Value

    End If
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    timWait.Enabled = False
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0

    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0

    Exit Sub
End If
mod1.Ti = mod1.Ti + 1
mod1.WP.Close
Set mod1.WP = Nothing
timWait.Enabled = True
End Sub

Private Sub txtOriginallyNumbers_DblClick()
frmRealNumbers.Visible = True
For oo = 0 To Me.dtgRealNumbers.Rows - 1
    Me.dtgRealNumbers.Row = oo
    For ii = 0 To 5
        Me.dtgRealNumbers.Col = ii
        dtgRealNumbers.CellForeColor = Me.txtOriginallyNumbers.ForeColor
    Next
Next
lblB1.Caption = "原厂零件编号:"
lblNumbers.Caption = txtOriginallyNumbers.Text
lblB1.ForeColor = txtOriginallyNumbers.ForeColor
lblNumbers.ForeColor = Me.txtOriginallyNumbers.ForeColor
End Sub


Private Sub txtPartName_Change()
txtPinYin.Text = UCase(GetCode(txtPartName.Text))
End Sub

Public Sub Initialize()

Me.txtHMNumbers.Text = ""


Me.txtPartName.Text = ""
Me.txtPartsCategory1.Text = ""
Me.txtPartsCategory2.Text = ""
Me.txtPinYin.Text = ""
Me.txtReplaceNumber1.Text = ""
Me.txtReplaceNumber2.Text = ""
Me.txtUnitBrand.Text = ""
Me.txtUnitModel.Text = ""
Me.txtUnitSeries.Text = ""


Me.dtgRealNumbers.Clear
Me.dtgRealNumbers.Cols = 6
Me.dtgRealNumbers.Row = 0: Me.dtgRealNumbers.Col = 1
Me.dtgRealNumbers.Text = "实物型号": Me.dtgRealNumbers.CellFontBold = True
Me.dtgRealNumbers.Col = 2: Me.dtgRealNumbers.Text = "描述": Me.dtgRealNumbers.CellFontBold = True
Me.dtgRealNumbers.Col = 3: Me.dtgRealNumbers.Text = "面价": Me.dtgRealNumbers.CellFontBold = True
Me.dtgRealNumbers.Col = 4: Me.dtgRealNumbers.Text = "基准价": Me.dtgRealNumbers.CellFontBold = True
Me.dtgRealNumbers.ColWidth(0) = 0
Me.dtgRealNumbers.ColWidth(1) = 2040
Me.dtgRealNumbers.ColWidth(2) = 2340
Me.dtgRealNumbers.ColWidth(3) = 885
Me.dtgRealNumbers.ColWidth(4) = 885
Me.dtgRealNumbers.ColWidth(5) = 0

Me.dtgSupplier.Cols = 4
Me.dtgSupplier.Clear
Me.dtgSupplier.Row = 0: Me.dtgSupplier.Col = 1
Me.dtgSupplier.Text = "供应商名称": Me.dtgSupplier.CellFontBold = True
Me.dtgSupplier.Col = 2: Me.dtgSupplier.Text = "联系人": Me.dtgSupplier.CellFontBold = True
Me.dtgSupplier.Col = 3: Me.dtgSupplier.Text = "最低成交价": Me.dtgSupplier.CellFontBold = True
Me.dtgSupplier.ColWidth(0) = 0
Me.dtgSupplier.ColWidth(1) = 2895
Me.dtgSupplier.ColWidth(2) = 1410
Me.dtgSupplier.ColWidth(3) = 1245

Me.lblRealPart.Caption = ""
Me.txtHMNumbers.ToolTipText = ""
Me.lblPP.Caption = ""
frmRealNumbers.Visible = False
End Sub

Private Sub txtReplaceNumber1_DblClick()
Dim oo As Integer
Dim ii As Integer
For oo = 0 To Me.dtgRealNumbers.Rows - 1
    Me.dtgRealNumbers.Row = oo
    For ii = 0 To 5
        Me.dtgRealNumbers.Col = ii
        dtgRealNumbers.CellForeColor = Me.txtReplaceNumber1.ForeColor
    Next
Next
frmRealNumbers.Visible = True
lblB1.Caption = "渠道替代编号:"
lblNumbers.Caption = txtReplaceNumber1.Text
lblB1.ForeColor = txtReplaceNumber1.ForeColor
lblNumbers.ForeColor = Me.txtReplaceNumber1.ForeColor
End Sub


Private Sub txtReplaceNumber2_Click()
Dim oo As Integer
Dim ii As Integer
For oo = 0 To Me.dtgRealNumbers.Rows - 1
    Me.dtgRealNumbers.Row = oo
    For ii = 0 To 5
        Me.dtgRealNumbers.Col = ii
        dtgRealNumbers.CellForeColor = Me.txtReplaceNumber2.ForeColor
    Next
Next
frmRealNumbers.Visible = True
lblB1.Caption = "功能替代编号"
lblNumbers.Caption = txtReplaceNumber2.Text
lblB1.ForeColor = txtReplaceNumber2.ForeColor
lblNumbers.ForeColor = Me.txtReplaceNumber2.ForeColor
End Sub



Public Sub BoundHM(Hid As Long)
Dim tt As String
Dim Ra
tt = "select unitBrand,unitSeries,unitModel,partsCategory1,partsCategory2,partName,pinyin,HMNumbers,originallyNumbers,replaceNumber1,replaceNumber2,PP from NLPG " & _
    " where hid=" & Hid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
'Call Me.initialize
Me.txtUnitBrand = Ra(0, 0)
Me.txtUnitSeries = Ra(1, 0)
Me.txtUnitModel = Ra(2, 0)
Me.txtPartsCategory1 = Ra(3, 0)
Me.txtPartsCategory2 = Ra(4, 0)
Me.txtPartName = Ra(5, 0)
Me.txtPinYin = Ra(6, 0)
Me.txtHMNumbers = Ra(7, 0)
Me.txtOriginallyNumbers = Ra(8, 0)
Me.txtReplaceNumber1 = Ra(9, 0)
Me.txtReplaceNumber2 = Ra(10, 0)
Me.txtHMNumbers.ToolTipText = Hid
Me.lblPP.Caption = Ra(11, 0)
End Sub
