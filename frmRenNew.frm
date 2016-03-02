VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form frmRenNew 
   BackColor       =   &H00C0FFC0&
   Caption         =   "公司组织架构"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8820
   ScaleWidth      =   15210
   Tag             =   "1"
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "运维部"
      Height          =   1185
      Index           =   14
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6120
      Width           =   465
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "项目部"
      Height          =   1185
      Index           =   13
      Left            =   13920
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6120
      Width           =   465
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "运维管理部"
      Height          =   1185
      Index           =   8
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6120
      Width           =   465
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "副总经理"
      Height          =   525
      Index           =   12
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3360
      Width           =   1275
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "业务二部"
      Height          =   705
      Index           =   11
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "副总经理"
      Height          =   525
      Index           =   10
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3360
      Width           =   1275
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "  市场部(业务四部)"
      Height          =   705
      Index           =   2
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4920
      Width           =   1065
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "商务部"
      Height          =   585
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4080
      Width           =   1665
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "总工"
      Height          =   585
      Index           =   6
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3360
      Width           =   1305
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FF8080&
      Caption         =   "运行五部"
      Height          =   375
      Index           =   30
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "钱庆 陶家强 王兵 贺德聪 孟友文 施加银 孙红来 张心 杜文忠 杜卫奇 沈彬彬 杨桃军 水海生 水春发 董宽 穆怀志"
      Top             =   9240
      Width           =   885
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FF8080&
      Caption         =   "运行四部"
      Height          =   375
      Index           =   29
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "赵克平 卓勤芳 黄建斌"
      Top             =   8790
      Width           =   885
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FF8080&
      Caption         =   "运行三部"
      Height          =   375
      Index           =   28
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "方永明"
      Top             =   8790
      Width           =   885
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FF8080&
      Caption         =   "运行二部"
      Height          =   375
      Index           =   27
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "袁金安 赵宝营 李午阳 罗俊斌 栗晓雷 龚进东 张海明 刘程 盛建华"
      Top             =   8820
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FF8080&
      Caption         =   "运行一部"
      Height          =   375
      Index           =   26
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "方永明 杨昌勋 朱彬 唐涛 陆炳康 汪协靖 汤成勋 袁露露 季政岗 陈中 周桂珍 严玉芳 胡爱周 冯敏 杨震国 陈华民"
      Top             =   8820
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "副总经理"
      Height          =   525
      Index           =   16
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3360
      Width           =   1275
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "  研发部(业务一部)"
      Height          =   705
      Index           =   15
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4920
      Width           =   1035
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "业务三部"
      Height          =   705
      Index           =   9
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4920
      Width           =   1065
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "行政人事部"
      Height          =   585
      Index           =   7
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   1545
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "工程管理部"
      Height          =   705
      Index           =   5
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1665
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "外地办事处"
      Height          =   705
      Index           =   3
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   1305
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "总经理"
      Height          =   585
      Index           =   1
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "总经理"
      Top             =   1350
      Width           =   1995
   End
   Begin VB.CommandButton cmdBM 
      BackColor       =   &H00FFFFC0&
      Caption         =   "董事会"
      Height          =   585
      Index           =   0
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   330
      Width           =   1995
   End
   Begin VB.Frame frmDetail 
      BackColor       =   &H00FFFFC0&
      Height          =   2925
      Left            =   1320
      TabIndex        =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   4095
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgA 
         Height          =   2925
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5159
         _Version        =   393216
         BackColor       =   12648384
         Rows            =   10
         Cols            =   3
         FixedCols       =   0
         BackColorFixed  =   16777152
         BackColorBkg    =   12648384
         BackColorUnpopulated=   8454016
         GridColorUnpopulated=   8454016
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
   End
   Begin NiceFormControl.NiceForm NF 
      Left            =   690
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton cmdOld 
      BackColor       =   &H00C0E0FF&
      Caption         =   "查看公司信息"
      Height          =   285
      Left            =   8220
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7950
      Width           =   1335
   End
   Begin VB.CommandButton cmdR 
      BackColor       =   &H00C0FFC0&
      Caption         =   "搜索"
      Height          =   285
      Left            =   11520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7980
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtR 
      Height          =   285
      Left            =   10560
      TabIndex        =   3
      Top             =   7950
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14430
      Picture         =   "frmRenNew.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   675
   End
   Begin VB.Line Line19 
      X1              =   14160
      X2              =   14160
      Y1              =   6720
      Y2              =   5280
   End
   Begin VB.Line Line17 
      X1              =   13320
      X2              =   13320
      Y1              =   6480
      Y2              =   5880
   End
   Begin VB.Line Line10 
      X1              =   15000
      X2              =   15000
      Y1              =   6480
      Y2              =   5880
   End
   Begin VB.Line Line9 
      X1              =   13320
      X2              =   15000
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line8 
      X1              =   13560
      X2              =   13560
      Y1              =   3720
      Y2              =   5280
   End
   Begin VB.Line Line6 
      X1              =   12480
      X2              =   12480
      Y1              =   3840
      Y2              =   5520
   End
   Begin VB.Line Line5 
      X1              =   6480
      X2              =   6480
      Y1              =   2640
      Y2              =   5280
   End
   Begin VB.Line Line4 
      X1              =   4920
      X2              =   4920
      Y1              =   2640
      Y2              =   5160
   End
   Begin VB.Line Line3 
      X1              =   3000
      X2              =   3000
      Y1              =   2640
      Y2              =   4320
   End
   Begin VB.Line Line18 
      X1              =   13080
      X2              =   13080
      Y1              =   3600
      Y2              =   2640
   End
   Begin VB.Line Line16 
      X1              =   10680
      X2              =   10680
      Y1              =   5400
      Y2              =   2640
   End
   Begin VB.Line Line12 
      X1              =   8520
      X2              =   8520
      Y1              =   5280
      Y2              =   2640
   End
   Begin VB.Line Line11 
      X1              =   840
      X2              =   840
      Y1              =   4440
      Y2              =   2640
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   13080
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line7 
      X1              =   7080
      X2              =   7080
      Y1              =   1800
      Y2              =   2640
   End
   Begin VB.Line Line15 
      X1              =   6000
      X2              =   5310
      Y1              =   9510
      Y2              =   9510
   End
   Begin VB.Line Line14 
      X1              =   5340
      X2              =   7050
      Y1              =   8970
      Y2              =   8970
   End
   Begin VB.Line Line13 
      Visible         =   0   'False
      X1              =   5340
      X2              =   6990
      Y1              =   8970
      Y2              =   9000
   End
   Begin VB.Line Line2 
      X1              =   7080
      X2              =   7080
      Y1              =   540
      Y2              =   1710
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "没有找到"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   12270
      TabIndex        =   7
      Top             =   8010
      Width           =   1725
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "人员查询"
      Height          =   285
      Left            =   9600
      TabIndex        =   6
      Top             =   8010
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "frmRenNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RRa
Dim LLa As Integer
Dim LColor As Long

Private Sub cmdBack_Click()
Me.Visible = False
frmZu.TBa.Buttons(7).Value = tbrUnpressed
End Sub

Private Sub cmdBm_Click(Index As Integer)
Dim oo As Integer
Dim ii As Integer
Exit Sub
If cmdBM(Index).Caption = frmDetail.ToolTipText Then
    If frmDetail.Visible = True Then
        frmDetail.Visible = False
        Exit Sub
    End If
End If
frmDetail.Left = cmdBM(Index).Left
frmDetail.Top = cmdBM(Index).Top + cmdBM(Index).Height
If frmDetail.Left + frmDetail.Width > Me.Width Then
    frmDetail.Left = Me.Width - frmDetail.Width
End If
If frmDetail.Top + frmDetail.Height > Me.Height Then
    'frmDetail.Top = Me.Height - frmDetail.Height
    frmDetail.Top = cmdBM(Index).Top - frmDetail.Height
End If
frmDetail.ToolTipText = cmdBM(Index).Caption
'设置人员明细
For oo = 0 To LLa
    If RRa(0, oo) = cmdBM(Index).Tag Then
        Exit For
    End If
Next
ii = 1: dtgA.Clear
dtgA.Row = 0
dtgA.Col = 0: dtgA.Text = "姓名": dtgA.CellFontBold = True
dtgA.Col = 1: dtgA.Text = "职务": dtgA.CellFontBold = True
dtgA.Col = 2: dtgA.Text = "联系方式": dtgA.CellFontBold = True
For oo = oo To LLa
    If RRa(0, oo) = cmdBM(Index).Tag Then
        dtgA.Row = ii
        dtgA.Col = 0: dtgA.Text = RRa(2, oo)
        dtgA.Col = 1: dtgA.Text = RRa(1, oo)
        dtgA.Col = 2: dtgA.Text = RRa(3, oo)
        ii = ii + 1
    Else
        Exit For
    End If
Next
frmDetail.Visible = True
End Sub

Private Sub cmdOld_Click()
DHB.Show
End Sub

Private Sub cmdR_Click()
Dim oo As Integer
Dim ZD As Boolean

For oo = 0 To 30
    If InStr(1, cmdBM(oo).ToolTipText, txtR.Text) > 0 Then
        cmdBM(oo).BackColor = &HFF&
    Else
        cmdBM(oo).BackColor = cmdBM(oo).Tag
    End If
Next
End Sub

Private Sub Form_Click()
frmDetail.Visible = False
lblTitle.Visible = False
End Sub

Private Sub Form_DblClick()
If mod1.DName = "马晓聪" Or mod1.DName = "陈珊珊" Then
    DHB.Show
End If
End Sub


Private Sub Form_Load()

Dim tt As String
Dim oo As Integer: Dim ii As Integer

NF.LoadSkin 3
Me.Height = 9330
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
dtgA.Rows = 30
dtgA.ColWidth(0) = 900
dtgA.ColWidth(1) = 1200
dtgA.ColWidth(2) = 1660
Set mod1.HTP = CreateObject("adodb.recordset")
'tt = "select bmid,userzw,username,userpho+' '+phoX from worker where zzf=1 order by bmid"
''''''''''tt = "SELECT dbo.BM.BMID, dbo.worker.UserZw, dbo.worker.UserName, dbo.worker.UserPho + ' ' + dbo.worker.phoX FROM dbo.worker INNER JOIN" & _
''''''''''      " dbo.BM ON dbo.worker.BM = dbo.BM.BM Where (dbo.worker.ZZF = 1) ORDER BY dbo.worker.BM"
''''''''''mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
''''''''''RRa = mod1.HTP.GetRows
''''''''''LLa = UBound(RRa, 2)
''''''''''On Error Resume Next








For oo = 0 To 30
    On Error Resume Next
    cmdBM(oo).Tag = cmdBM(oo).BackColor

Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
frmZu.TBa.Buttons(7).Value = tbrUnpressed
Cancel = True
End Sub

Private Sub txtR_KeyDown(KeyCode As Integer, Shift As Integer)
Dim oo As Integer
Dim ZD As Boolean

If KeyCode = 13 Then
    For oo = 0 To 24
    LColor = cmdBM(oo).BackColor
    If InStr(1, cmdBM(oo).ToolTipText, txtR.Text) > 0 Then
        cmdBM(oo).BackColor = &H8080FF
        ZD = True
    Else
        cmdBM(oo).BackColor = LColor
    End If
    Next
End If
If ZD = False Then
    lblTitle.Visible = True
End If
End Sub


