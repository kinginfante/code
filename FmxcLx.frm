VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form FmxcLx 
   BackColor       =   &H00C0FFFF&
   Caption         =   "业务类型"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   7605
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2910
      Top             =   4110
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   4050
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLx 
      Height          =   4065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7170
      _Version        =   393216
      BackColor       =   12648447
      Rows            =   14
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   16777152
      BackColorBkg    =   12648447
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin NiceFormControl.NiceButton cmdNew 
      Height          =   345
      Left            =   4170
      TabIndex        =   2
      Top             =   4200
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   609
      BTYPE           =   3
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FmxcLx.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
      Caption         =   "生成询价单"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请用鼠标双击相应的栏目"
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   4260
      Width           =   2535
   End
End
Attribute VB_Name = "FmxcLx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LX As String
Private Sub cmdNew_Click()
If cmdNew.Caption = "生成询价单" Then
    MsgBox "请选择相应的业务类型!(在列表中双击)"
    Exit Sub
End If
Call FMXCXmmc.Qing
FMXCXmmc.Show
FMXCXmmc.ZOrder 0
FMXCXmmc.Lb = "询价单"
FMXCXmmc.NiceButton1.Caption = "生 成 单 据 (询价单)"
End Sub

Private Sub dtgLx_DblClick()

If dtgLx.Row = 0 Then Exit Sub
dtgLx.Col = 1: LX = dtgLx.Text
cmdNew.Caption = "生成" & LX & "询价单"
cmdNew.ToolTipText = dtgLx.Row

End Sub

Private Sub Form_Load()
Me.Width = 7725
Me.Height = 5190

dtgLx.Row = 0
dtgLx.Col = 0: dtgLx.Text = "业务类型"
dtgLx.Col = 1: dtgLx.Text = "业务类型"
dtgLx.Col = 2: dtgLx.Text = "基准价格"
dtgLx.Col = 3: dtgLx.Text = "速达金额"
dtgLx.Col = 4: dtgLx.Text = "询价单"
dtgLx.Col = 5: dtgLx.Text = "合同编号"
dtgLx.Col = 6: dtgLx.Text = "说明"
dtgLx.MergeCells = flexMergeFree
dtgLx.MergeRow(0) = True
dtgLx.Row = 1: dtgLx.Col = 0: dtgLx.Text = "人工类"
dtgLx.Row = 1: dtgLx.Col = 1: dtgLx.Text = "维保": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 1: dtgLx.Col = 6: dtgLx.Text = "本公司人员自行完成的人工"
dtgLx.Row = 2: dtgLx.Col = 0: dtgLx.Text = "人工类"
dtgLx.Row = 2: dtgLx.Col = 1: dtgLx.Text = "大修": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 2: dtgLx.Col = 6: dtgLx.Text = "本公司人员自行完成的人工"
dtgLx.Row = 3: dtgLx.Col = 0: dtgLx.Text = "人工类"
dtgLx.Row = 3: dtgLx.Col = 1: dtgLx.Text = "其他人工": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 3: dtgLx.Col = 6: dtgLx.Text = "本公司人员自行完成的人工"
dtgLx.Row = 4: dtgLx.Col = 0: dtgLx.Text = "压缩机"
dtgLx.Row = 4: dtgLx.Col = 1: dtgLx.Text = "压缩机维修保养": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 4: dtgLx.Col = 6: dtgLx.Text = "压缩机工厂的维修或保养"
dtgLx.Row = 5: dtgLx.Col = 0: dtgLx.Text = "压缩机"
dtgLx.Row = 5: dtgLx.Col = 1: dtgLx.Text = "压缩机贸易": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 5: dtgLx.Col = 6: dtgLx.Text = "压缩机工厂的产品销售"
dtgLx.Row = 6: dtgLx.Col = 0: dtgLx.Text = "中介"
dtgLx.Row = 6: dtgLx.Col = 1: dtgLx.Text = "中介业务"
dtgLx.Row = 6: dtgLx.Col = 6: dtgLx.Text = "中介（居间）业务收入"
dtgLx.Row = 7: dtgLx.Col = 0: dtgLx.Text = "贸易"
dtgLx.Row = 7: dtgLx.Col = 1: dtgLx.Text = "三菱": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 7: dtgLx.Col = 6: dtgLx.Text = "三菱设备的贸易"
dtgLx.Row = 8: dtgLx.Col = 0: dtgLx.Text = "贸易"
dtgLx.Row = 8: dtgLx.Col = 1: dtgLx.Text = "松下": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 8: dtgLx.Col = 6: dtgLx.Text = "广州杰狮对外松下设备的贸易"
dtgLx.Row = 9: dtgLx.Col = 0: dtgLx.Text = "贸易"
dtgLx.Row = 9: dtgLx.Col = 1: dtgLx.Text = "勤达富": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 9: dtgLx.Col = 6: dtgLx.Text = "勤达富设备的贸易"
dtgLx.Row = 10: dtgLx.Col = 0: dtgLx.Text = "贸易"
dtgLx.Row = 10: dtgLx.Col = 1: dtgLx.Text = "德图": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 10: dtgLx.Col = 6: dtgLx.Text = "德图设备的贸易"
dtgLx.Row = 11: dtgLx.Col = 0: dtgLx.Text = "贸易"
dtgLx.Row = 11: dtgLx.Col = 1: dtgLx.Text = "零配件": dtgLx.CellBackColor = &HC0FFC0
dtgLx.Row = 11: dtgLx.Col = 6: dtgLx.Text = "零配件（包含工具易耗）的贸易"
dtgLx.Row = 12: dtgLx.Col = 0: dtgLx.Text = "贸易"
dtgLx.Row = 12: dtgLx.Col = 1: dtgLx.Text = "分包"
dtgLx.Row = 12: dtgLx.Col = 6: dtgLx.Text = "分包合同"
dtgLx.Row = 13: dtgLx.Col = 0: dtgLx.Text = "贸易"
dtgLx.Row = 13: dtgLx.Col = 1: dtgLx.Text = "非代理产品"
dtgLx.Row = 13: dtgLx.Col = 6: dtgLx.Text = "非代理产品的贸易"
dtgLx.Col = 5
dtgLx.Row = 1: dtgLx.Text = "RG": dtgLx.Row = 2: dtgLx.Text = "RG": dtgLx.Row = 3: dtgLx.Text = "RG"
dtgLx.Row = 4: dtgLx.Text = "YS": dtgLx.Row = 5: dtgLx.Text = "YS"
dtgLx.Row = 6: dtgLx.Text = "ZJ"
dtgLx.Row = 7: dtgLx.Text = "TR": dtgLx.Row = 8: dtgLx.Text = "TR": dtgLx.Row = 8: dtgLx.Text = "TR": dtgLx.Row = 10: dtgLx.Text = "TR": dtgLx.Row = 11: dtgLx.Text = "TR"
dtgLx.Row = 12: dtgLx.Text = "TR": dtgLx.Row = 13: dtgLx.Text = "TR": dtgLx.Row = 9: dtgLx.Text = "TR"
dtgLx.MergeCol(5) = True
dtgLx.MergeCol(0) = True
dtgLx.ColWidth(1) = 1695
dtgLx.ColWidth(2) = 0
dtgLx.ColWidth(3) = 0
dtgLx.ColWidth(4) = 0
dtgLx.ColWidth(6) = 2925
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
Cancel = True
End Sub


