VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmTDCG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "采购信息"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmGy 
      Caption         =   "供应商资料"
      ForeColor       =   &H00008000&
      Height          =   4815
      Left            =   1500
      TabIndex        =   21
      Top             =   3240
      Width           =   5085
      Begin VB.Frame frmQm 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   1155
         Left            =   120
         TabIndex        =   67
         Top             =   3600
         Width           =   4575
         Begin VB.CommandButton cmdQm 
            Caption         =   "cmdQm"
            Height          =   345
            Index           =   1
            Left            =   1530
            TabIndex        =   77
            Top             =   330
            Width           =   945
         End
         Begin VB.CommandButton cmdQm 
            Caption         =   "cmdQm"
            Height          =   345
            Index           =   0
            Left            =   480
            TabIndex        =   69
            Top             =   330
            Width           =   945
         End
         Begin VB.CommandButton cmdPje 
            Caption         =   "评审建议"
            Height          =   1065
            Left            =   60
            TabIndex        =   68
            Top             =   90
            Width           =   345
         End
         Begin VB.Label lblFwid 
            Caption         =   "lblFwid"
            Height          =   255
            Left            =   3720
            TabIndex        =   83
            Top             =   720
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lblLcUid 
            Caption         =   "lblLcUid"
            Height          =   285
            Left            =   2670
            TabIndex        =   82
            Top             =   720
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label lblLcRen 
            Caption         =   "lblLcRen"
            Height          =   285
            Left            =   2700
            TabIndex        =   81
            Top             =   180
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lblLc 
            Caption         =   "lblLc"
            Height          =   315
            Left            =   3750
            TabIndex        =   80
            Top             =   210
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label lblTm 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   1
            Left            =   1530
            TabIndex        =   79
            Top             =   750
            Width           =   945
         End
         Begin VB.Label lblQM 
            Caption         =   "组长"
            Height          =   195
            Index           =   1
            Left            =   1620
            TabIndex        =   78
            Top             =   90
            Width           =   885
         End
         Begin VB.Label lblQM 
            Caption         =   "录入者"
            Height          =   195
            Index           =   0
            Left            =   570
            TabIndex        =   71
            Top             =   90
            Width           =   885
         End
         Begin VB.Label lblTm 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   405
            Index           =   0
            Left            =   480
            TabIndex        =   70
            Top             =   750
            Width           =   945
         End
      End
      Begin VB.TextBox txtBz 
         ForeColor       =   &H0000C000&
         Height          =   660
         Left            =   1260
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   2910
         Width           =   3375
      End
      Begin VB.TextBox txtT2 
         ForeColor       =   &H0000C000&
         Height          =   270
         Left            =   1260
         TabIndex        =   34
         Top             =   2490
         Width           =   3375
      End
      Begin VB.TextBox txtLcRen2 
         ForeColor       =   &H0000C000&
         Height          =   270
         Left            =   1260
         TabIndex        =   33
         Top             =   2058
         Width           =   3375
      End
      Begin VB.TextBox txtT1 
         ForeColor       =   &H0000C000&
         Height          =   270
         Left            =   1260
         TabIndex        =   32
         Top             =   1626
         Width           =   3375
      End
      Begin VB.TextBox txtLcRen1 
         ForeColor       =   &H0000C000&
         Height          =   270
         Left            =   1260
         TabIndex        =   31
         Top             =   1194
         Width           =   3375
      End
      Begin VB.TextBox txtGAdr 
         ForeColor       =   &H0000C000&
         Height          =   270
         Left            =   1260
         TabIndex        =   30
         Top             =   762
         Width           =   3375
      End
      Begin VB.TextBox txtGmc 
         ForeColor       =   &H0000C000&
         Height          =   270
         Left            =   1260
         TabIndex        =   29
         Top             =   330
         Width           =   3375
      End
      Begin VB.Label lblGyid 
         Caption         =   "lblGyid"
         Height          =   195
         Left            =   210
         TabIndex        =   40
         Top             =   3360
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label20 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   165
         Index           =   3
         Left            =   1110
         TabIndex        =   39
         Top             =   1800
         Width           =   75
      End
      Begin VB.Label Label20 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   165
         Index           =   2
         Left            =   1110
         TabIndex        =   38
         Top             =   1320
         Width           =   75
      End
      Begin VB.Label Label20 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   165
         Index           =   1
         Left            =   1110
         TabIndex        =   37
         Top             =   840
         Width           =   75
      End
      Begin VB.Label Label20 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   165
         Index           =   0
         Left            =   1110
         TabIndex        =   36
         Top             =   360
         Width           =   75
      End
      Begin VB.Label Label18 
         Caption         =   "备注"
         ForeColor       =   &H0000C000&
         Height          =   225
         Left            =   180
         TabIndex        =   28
         Top             =   2970
         Width           =   1065
      End
      Begin VB.Label Label17 
         Caption         =   "联系2电话"
         ForeColor       =   &H0000C000&
         Height          =   225
         Left            =   180
         TabIndex        =   27
         Top             =   2535
         Width           =   1065
      End
      Begin VB.Label Label16 
         Caption         =   "联系人2"
         ForeColor       =   &H0000C000&
         Height          =   225
         Left            =   180
         TabIndex        =   26
         Top             =   2100
         Width           =   1065
      End
      Begin VB.Label Label15 
         Caption         =   "联系1电话"
         ForeColor       =   &H0000C000&
         Height          =   225
         Left            =   180
         TabIndex        =   25
         Top             =   1665
         Width           =   825
      End
      Begin VB.Label Label14 
         Caption         =   "联系人1"
         ForeColor       =   &H0000C000&
         Height          =   225
         Left            =   180
         TabIndex        =   24
         Top             =   1230
         Width           =   825
      End
      Begin VB.Label Label13 
         Caption         =   "供应商地址"
         ForeColor       =   &H0000C000&
         Height          =   225
         Left            =   180
         TabIndex        =   23
         Top             =   795
         Width           =   915
      End
      Begin VB.Label Label12 
         Caption         =   "供应商名称"
         ForeColor       =   &H0000C000&
         Height          =   225
         Left            =   180
         TabIndex        =   22
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4785
      Left            =   30
      TabIndex        =   72
      Top             =   4500
      Width           =   4305
      Begin VB.CommandButton cmdPr 
         Caption         =   "上一月"
         Height          =   315
         Left            =   390
         TabIndex        =   74
         Top             =   4410
         Width           =   1155
      End
      Begin VB.CommandButton cmdNe 
         Caption         =   "下一月"
         Height          =   315
         Left            =   2250
         TabIndex        =   73
         Top             =   4410
         Width           =   1155
      End
      Begin MSComCtl2.MonthView monDate 
         Height          =   2220
         Left            =   120
         TabIndex        =   75
         Top             =   2130
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   3916
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   12189625
         Appearance      =   1
         MonthBackColor  =   12189625
         ScrollRate      =   21
         StartOfWeek     =   57933825
         TitleBackColor  =   12615680
         TrailingForeColor=   -2147483635
         CurrentDate     =   39090
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "导航"
      Height          =   585
      Left            =   14550
      Picture         =   "frmTDCG.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8640
      Width           =   675
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgCGj 
      Height          =   4425
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   7805
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame2 
      Caption         =   "替换内容"
      ForeColor       =   &H000000FF&
      Height          =   4815
      Left            =   9390
      TabIndex        =   41
      Top             =   4470
      Width           =   5835
      Begin VB.Frame frmMod 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   915
         Left            =   180
         TabIndex        =   61
         Top             =   3690
         Width           =   4785
         Begin VB.TextBox txtJ 
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   10
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   64
            Top             =   60
            Width           =   3645
         End
         Begin VB.CommandButton cmdSave 
            Height          =   435
            Left            =   4290
            Picture         =   "frmTDCG.frx":0102
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   480
            Width           =   465
         End
         Begin VB.TextBox Text2 
            ForeColor       =   &H0000C000&
            Height          =   315
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   62
            Text            =   "双击查看"
            Top             =   450
            Width           =   2625
         End
         Begin VB.Label Label20 
            Caption         =   "*"
            ForeColor       =   &H000000FF&
            Height          =   165
            Index           =   6
            Left            =   840
            TabIndex        =   86
            Top             =   120
            Width           =   75
         End
         Begin VB.Label Label10 
            Caption         =   "提供者"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   0
            TabIndex        =   66
            Top             =   120
            Width           =   585
         End
         Begin VB.Label Label11 
            Caption         =   "供应商资料"
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   0
            TabIndex        =   65
            Top             =   510
            Width           =   945
         End
      End
      Begin VB.TextBox txtJ 
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   21
         Left            =   1260
         TabIndex        =   50
         Top             =   270
         Width           =   3645
      End
      Begin VB.TextBox txtJ 
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   20
         Left            =   1260
         TabIndex        =   49
         Top             =   647
         Width           =   3645
      End
      Begin VB.TextBox txtJ 
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   19
         Left            =   1260
         TabIndex        =   48
         Top             =   1024
         Width           =   3645
      End
      Begin VB.TextBox txtJ 
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   18
         Left            =   1260
         TabIndex        =   47
         Top             =   1401
         Width           =   3645
      End
      Begin VB.TextBox txtJ 
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   17
         Left            =   1260
         TabIndex        =   46
         Top             =   1778
         Width           =   3645
      End
      Begin VB.TextBox txtJ 
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   16
         Left            =   1260
         TabIndex        =   45
         Top             =   2155
         Width           =   3645
      End
      Begin VB.TextBox txtJ 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   15
         Left            =   1260
         TabIndex        =   44
         Top             =   2532
         Width           =   3645
      End
      Begin VB.TextBox txtJ 
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   14
         Left            =   1260
         TabIndex        =   43
         Top             =   2910
         Width           =   3645
      End
      Begin VB.TextBox txtJ 
         ForeColor       =   &H00C000C0&
         Height          =   270
         Index           =   9
         Left            =   1260
         TabIndex        =   42
         Top             =   3330
         Width           =   3645
      End
      Begin VB.Label Label20 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   165
         Index           =   4
         Left            =   1020
         TabIndex        =   85
         Top             =   2610
         Width           =   75
      End
      Begin VB.Label Label32 
         Caption         =   "机组品牌"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   180
         TabIndex        =   60
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label31 
         Caption         =   "机组型号"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   180
         TabIndex        =   59
         Top             =   685
         Width           =   915
      End
      Begin VB.Label Label30 
         Caption         =   "压缩机型号"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   180
         TabIndex        =   58
         Top             =   1070
         Width           =   945
      End
      Begin VB.Label Label29 
         Caption         =   "出厂编号"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   180
         TabIndex        =   57
         Top             =   1455
         Width           =   1005
      End
      Begin VB.Label Label28 
         Caption         =   "机组序列号"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   180
         TabIndex        =   56
         Top             =   1840
         Width           =   1005
      End
      Begin VB.Label Label27 
         Caption         =   "零件编号"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   180
         TabIndex        =   55
         Top             =   2225
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "零件名称"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   54
         Top             =   2610
         Width           =   1065
      End
      Begin VB.Label Label25 
         Caption         =   "品牌产地"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   180
         TabIndex        =   53
         Top             =   3002
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "替换价"
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   180
         TabIndex        =   52
         Top             =   3390
         Width           =   765
      End
      Begin VB.Label Label20 
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   165
         Index           =   5
         Left            =   1020
         TabIndex        =   51
         Top             =   3390
         Width           =   75
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "采购详情"
      Height          =   4785
      Left            =   4350
      TabIndex        =   2
      Top             =   4500
      Width           =   5025
      Begin VB.TextBox txtJ 
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   8
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   3330
         Width           =   3555
      End
      Begin VB.TextBox txtJ 
         Height          =   270
         Index           =   7
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2944
         Width           =   3555
      End
      Begin VB.TextBox txtJ 
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   6
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2562
         Width           =   3555
      End
      Begin VB.TextBox txtJ 
         Height          =   270
         Index           =   5
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2180
         Width           =   3555
      End
      Begin VB.TextBox txtJ 
         Height          =   270
         Index           =   4
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1798
         Width           =   3555
      End
      Begin VB.TextBox txtJ 
         Height          =   270
         Index           =   3
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1416
         Width           =   3555
      End
      Begin VB.TextBox txtJ 
         Height          =   270
         Index           =   2
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1034
         Width           =   3555
      End
      Begin VB.TextBox txtJ 
         Height          =   270
         Index           =   1
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   652
         Width           =   3555
      End
      Begin VB.TextBox txtJ 
         Height          =   270
         Index           =   0
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   3555
      End
      Begin VB.Label lblLId 
         Caption         =   "lblLId"
         Height          =   285
         Left            =   2280
         TabIndex        =   84
         Top             =   3960
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblTDid 
         Caption         =   "lblTDid"
         Height          =   315
         Left            =   480
         TabIndex        =   76
         Top             =   3930
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label9 
         Caption         =   "单价"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   3390
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "品牌产地"
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   3002
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "零件名称"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   2616
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "零件编号"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   2230
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "机组序列号"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   1844
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "出厂编号"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   1458
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "压缩机型号"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   1072
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "机组型号"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   686
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "机组品牌"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmTDCG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public adoGCJ As ADODB.Recordset
Public Drq As Date '当前日期
Public RenXuan As Integer
Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
frmZu.ZOrder 0
If Dialog.Visible = True Then
    Dialog.Enabled = True
    Dialog.ZOrder 0
End If
End Sub

Private Sub cmdNe_Click()
Dim tt As String
Dim Orq As Date
On Error Resume Next
Orq = DateSerial(Year(Drq), Month(Drq) + 2, 1)
Drq = DateSerial(Year(Drq), Month(Drq) + 1, 1)
tt = "select * from xunjiagcj where 询价日期>='" & Drq & "' and 询价日期<'" & Orq & "' order by 询价日期 desc"
adoGCJ.Close
adoGCJ.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If adoGCJ.RecordCount > 0 Then
    Set dtgCGj.DataSource = adoGCJ
    dtgCGj.FixedRows = 1
    dtgCGj.Row = adoGCJ.RecordCount - 1
Else
    'Set dtgCGj.DataSource = adoGCJ
    dtgCGj.Rows = 2
    dtgCGj.FixedRows = 1
    dtgCGj.Row = 1
    For oo = 0 To 10
        dtgCGj.Col = oo
        dtgCGj.Text = ""
    Next
End If
If monDate.Month < 12 Then
    monDate.Month = monDate.Month + 1
Else
    monDate.Month = 1
    monDate.Year = monDate.Year + 1
End If
If adoGCJ.RecordCount > 0 Then
dtgCGj.Row = adoGCJ.RecordCount - 1
End If
'adoGCJ.Close
End Sub

Private Sub cmdPr_Click()
Dim tt As String
Dim Orq As Date
On Error Resume Next
Orq = Drq
Drq = DateSerial(Year(Drq), Month(Drq) - 1, 1)
tt = "select * from xunjiagcj where 询价日期>='" & Drq & "' and 询价日期<'" & Orq & "' order by 询价日期 desc"
adoGCJ.Close
adoGCJ.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If adoGCJ.RecordCount > 0 Then
    Set dtgCGj.DataSource = adoGCJ
    dtgCGj.FixedRows = 1
    dtgCGj.Row = adoGCJ.RecordCount - 1
Else
    'Set dtgCGj.DataSource = adoGCJ
    dtgCGj.Rows = 2
    dtgCGj.FixedRows = 1
    dtgCGj.Row = 1
    For oo = 0 To 10
        dtgCGj.Col = oo
        dtgCGj.Text = ""
    Next
End If
If monDate.Month > 1 Then
    monDate.Month = monDate.Month - 1
Else
    monDate.Month = 12
    monDate.Year = monDate.Year - 1
End If
If adoGCJ.RecordCount > 0 Then
dtgCGj.Row = adoGCJ.RecordCount - 1
End If
'adoGCJ.Close
End Sub

Private Sub cmdQm_Click(Index As Integer)
Dim adoQm As ADODB.command
Dim tt As String

Dim ii As Integer
On Error Resume Next

If lblLc.Caption <> Index + 1 Then
    Exit Sub
End If
If mod1.DName <> lblLcRen.Caption Then
    MsgBox "此处应由" & lblLcRen.Caption & "签字，请不要乱点！"
    Exit Sub
End If

Set adoQm = New ADODB.command

If lblLc.Caption > 1 Then
    ii = MsgBox("您是否核准此单？(选择“是”将签字通过,选择“否”将驳回此单)", vbYesNoCancel + vbInformation, "请您注意!")
    If ii = vbNo Then
        ii = MsgBox("驳回后,此单将回转至填单人" & lblYwy.Caption & ",确认吗?", vbYesNo + vbInformation, "确认驳回吗?")
        If ii = vbNo Then
            Exit Sub
        End If
        tt = InputBox("请输入您要驳回的原因!")

        adoQm.ActiveConnection = mod1.CC
        adoQm.CommandText = "xtzxFAdd"
        adoQm.CommandType = adCmdStoredProc
        adoQm.Parameters("@yid").Value = 78  '反签名
        adoQm.Parameters("@lc").Value = lblLc.Caption
        adoQm.Parameters("@bh").Value = lblTDid.Caption
        adoQm.Parameters("@ywy").Value = mod1.DName
        adoQm.Parameters("@uid").Value = mod1.DHid
        adoQm.Parameters("@bz").Value = tt
        adoQm.Parameters("@zn").Value = lblQM(Index).Caption '身份职能
        adoQm.Execute
        Set adoQm = Nothing
        For oo = 0 To 6
            cmdQm(oo).Caption = ""
            lblTm(oo).Caption = ""
        Next
        lblLc.Caption = 999 '不让再按签名按钮.
        If Dialog.Visible = True Then '更新事务列表
            Call mod1.refEnvent(1)
        End If
        Exit Sub
    ElseIf ii = vbCancel Then
        Exit Sub
    End If
End If

ii = MsgBox("是否确认签字?", vbYesNo)
If ii = vbNo Then Exit Sub
tt = InputBox("如果您有评审建议,请在此输入!")

adoQm.ActiveConnection = mod1.CC
adoQm.CommandText = "QMtdp"
adoQm.CommandType = adCmdStoredProc
adoQm.Parameters("@lc").Value = Val(lblLc.Caption)
adoQm.Parameters("@ywy").Value = mod1.DName
adoQm.Parameters("@uid").Value = mod1.DHid
adoQm.Parameters("@bz").Value = tt
adoQm.Parameters("@fwid").Value = lblFwid.Caption
adoQm.Parameters("@bh").Value = Val(lblTDid.Caption)
adoQm.Parameters("@yid").Value = 78
adoQm.Parameters("@nr").Value = txtJ(15).Text
adoQm.Parameters("@lab").Value = lblQM(Index + 1).Caption
adoQm.Parameters("@dxren").Value = txtJ(10).Text
adoQm.Parameters("@dxuid").Value = txtJ(10).ToolTipText
adoQm.Parameters("@errch").Value = ""
adoQm.Parameters("@Tywy").Value = ""
adoQm.Parameters("@Tuid").Value = ""
adoQm.Execute
tt = adoQm.Parameters("@errch").Value
Tywy = adoQm.Parameters("@Tywy").Value
Tuid = adoQm.Parameters("@Tuid").Value
Set adoQm = Nothing

 
If tt = "成功" Then
    cmdQm(Index).Caption = mod1.DName
    lblTm(Index).Caption = mod1.DQda
    lblLcRen.Caption = Tywy
    lblLcUid.Caption = Tuid
    lblLc.Caption = lblLc.Caption + 1
    MsgBox tt

    
    frmBxBrow.adoYj.Requery
    Set frmBxBrow.dtgYJ.DataSource = frmBxBrow.adoYj
    
    If Dialog.Visible = True Then '更新事务列表
    Call mod1.refEnvent(1)
    End If
Else
        MsgBox "网络出现故障,请再试一次,如果还是提交不成功,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
End If
End Sub

Private Sub cmdQm_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ii As Integer
If Index = 0 Or Button = 1 Then Exit Sub
ii = MsgBox("是否要转移此单审核人?", vbInformation + vbYesNo, "您好!")
If ii = vbNo Then Exit Sub
Set Ren.XForm = New frmTDCG
Call mod1.RenXz("frmTDCG", Me, 1)
RenXuan = 1

End Sub

Private Sub cmdSave_Click()
Dim cmd As ADODB.command
Dim ERRch As String
On Error Resume Next
If Val(lblLc.Caption) > 1 Then
    MsgBox "请让录入者来填写!"
    Exit Sub
End If
Set cmd = New ADODB.command
cmd.ActiveConnection = mod1.CC
cmd.CommandText = "XTDadd"
cmd.CommandType = adCmdStoredProc
cmd.Parameters("@jzPb") = txtJ(21).Text
cmd.Parameters("@jzxh") = txtJ(20).Text
cmd.Parameters("@yXh") = txtJ(19).Text
cmd.Parameters("@CCbh") = txtJ(18).Text
cmd.Parameters("@jzBh") = txtJ(17).Text
cmd.Parameters("@ljBh") = txtJ(16).Text
cmd.Parameters("@ljMc") = txtJ(15).Text
cmd.Parameters("@pbcd") = txtJ(14).Text
cmd.Parameters("@Tdj") = txtJ(9).Text
cmd.Parameters("@tdRen") = txtJ(10).Text
cmd.Parameters("@uid") = txtJ(10).ToolTipText
cmd.Parameters("@TDid") = Val(lblTDid.Caption)
cmd.Parameters("@Rywy") = mod1.DName
cmd.Parameters("@Ruid") = mod1.DHid
cmd.Parameters("@errch") = ""
cmd.Parameters("@gmc") = txtGmc.Text
cmd.Parameters("@gAdr") = txtGAdr.Text
cmd.Parameters("@LcRen1") = txtLcRen1.Text
cmd.Parameters("@LcRen2") = txtLcRen2.Text
cmd.Parameters("@T1") = txtT1.Text
cmd.Parameters("@T2") = txtT2.Text
cmd.Parameters("@Bz") = txtBz.Text
cmd.Parameters("@gyid") = Val(lblGyid.Caption)
cmd.Parameters("@lid") = lblLId.Caption
cmd.Execute
ERRch = cmd.Parameters("@errch").Value
If ERRch <> "成功" Then
        MsgBox "网络出现故障,请再试一次,如果还是提交不成功,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        Exit Sub
End If
lblTDid.Caption = cmd.Parameters("@tdid").Value
lblGyid.Caption = cmd.Parameters("@gyid").Value

Set cmd = Nothing
'只有录入者,在流程张一部,才有权限保存
lblLc.Caption = 1
lblLcRen.Caption = mod1.DName
lblLcUid.Caption = mod1.DHid
MsgBox "OK!"
adoGCJ.Requery
Set Me.dtgCGj.DataSource = adoGCJ
End Sub

Private Sub dtgCGj_Click()
'Dim oo As Integer
On Error Resume Next
'Dim oo As Integer
'On Error Resume Next
Call TDQing
Call TDBound
'For oo = 0 To 8
'    txtJ(oo).Text = ""
'Next
'txtJ(0).Text = adoGCJ.Fields("机组品牌").Value
'txtJ(1).Text = adoGCJ.Fields("机组型号").Value
'txtJ(2).Text = adoGCJ.Fields("压缩机型号").Value
'txtJ(3).Text = adoGCJ.Fields("出厂编号").Value
'txtJ(4).Text = adoGCJ.Fields("机组序列号").Value
'txtJ(5).Text = adoGCJ.Fields("零件编号").Value
'txtJ(6).Text = adoGCJ.Fields("零件名称").Value
'txtJ(7).Text = adoGCJ.Fields("品牌产地").Value
'txtJ(8).Text = adoGCJ.Fields("单价").Value
End Sub

Private Sub dtgCGj_RowColChange()
'Dim oo As Integer
On Error Resume Next
Call TDQing
Call TDBound
'dtgCGj.Col = 1
'txtJ(0).Text = adoGCJ.Fields("机组品牌").Value
'txtJ(1).Text = adoGCJ.Fields("机组型号").Value
'txtJ(2).Text = adoGCJ.Fields("压缩机型号").Value
'txtJ(3).Text = adoGCJ.Fields("出厂编号").Value
'txtJ(4).Text = adoGCJ.Fields("机组序列号").Value
'txtJ(5).Text = adoGCJ.Fields("零件编号").Value
'txtJ(6).Text = adoGCJ.Fields("零件名称").Value
'txtJ(7).Text = adoGCJ.Fields("品牌产地").Value
'txtJ(8).Text = adoGCJ.Fields("单价").Value
End Sub

Private Sub Form_Load()

Set adoGCJ = New ADODB.Recordset
Me.Height = mod1.FHeight
Me.Width = mod1.FWidth
Me.Left = 0
Me.Top = 0
dtgCGj.ColWidth(0) = 300
dtgCGj.ColWidth(5) = 0
dtgCGj.ColWidth(6) = 0
dtgCGj.ColWidth(3) = 2000
dtgCGj.ColWidth(4) = 2000
dtgCGj.ColWidth(7) = 2500
dtgCGj.ColWidth(8) = 2500
dtgCGj.ColWidth(12) = 0
dtgCGj.ColWidth(13) = 700
dtgCGj.ColWidth(14) = 700 'TDID
If mod1.DName = "冯建川" Then
    frmMod.Visible = True
Else
    frmMod.Visible = False
End If
frmGy.Left = 4350
frmGy.Top = 4500
frmGy.Visible = False

End Sub






Private Sub monDate_DateDblClick(ByVal DateDblClicked As Date)
Dim tt As String
Dim Orq As Date
Dim oo As Integer
On Error Resume Next
'MsgBox DateDblClicked
tt = "select * from xunjiagcj where 询价日期>='" & monDate.Value & _
"' and 询价日期<'" & DateSerial(Year(monDate.Value), Month(monDate.Value), Day(monDate.Value) + 1) & "' order by 询价日期 desc"
adoGCJ.Close
adoGCJ.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If adoGCJ.RecordCount > 0 Then
    Set dtgCGj.DataSource = adoGCJ
    dtgCGj.FixedRows = 1
    dtgCGj.Row = adoGCJ.RecordCount - 1
Else
    'Set dtgCGj.DataSource = adoGCJ
    dtgCGj.Rows = 2
    dtgCGj.FixedRows = 1
    dtgCGj.Row = 1
    For oo = 0 To 10
        dtgCGj.Col = oo
        dtgCGj.Text = ""
    Next
End If
'If adoGCJ.RecordCount > 0 Then
'dtgCGj.Row = adoGCJ.RecordCount - 1
'End If

End Sub






Private Sub Text2_DblClick()
If frmGy.Visible = False Then
    frmGy.Visible = True
Else
    frmGy.Visible = False
End If

End Sub

Private Sub txtJ_DblClick(Index As Integer)
If Index = 10 Then
    Set Ren.XForm = New frmTDCG
    Call mod1.RenXz("frmTDCG", Me, 0)
    RenXuan = 0
End If
End Sub



Public Sub TDQing()
Dim oo As Integer
For oo = 0 To 10
    txtJ(oo).Text = ""
Next
txtJ(10).ToolTipText = ""
For oo = 14 To 21
    txtJ(oo).Text = ""
Next
lblGyid.Caption = ""
lblTDid.Caption = ""
txtGmc.Text = ""
txtGAdr.Text = ""
txtLcRen1.Text = ""
txtLcRen2.Text = ""
txtT1.Text = ""
txtT2.Text = ""
txtBz.Text = ""
cmdQm(0).Caption = ""
cmdQm(1).Caption = ""
lblTm(0).Caption = ""
lblTm(1).Caption = ""
lblLc.Caption = ""
lblFwid.Caption = ""
lblLcRen.Caption = ""
lblLcUid.Caption = ""
lblLId.Caption = ""
End Sub

Public Sub TDBound()
Dim oo As Integer
Dim tt As String

Dim adoTd As ADODB.Recordset
On Error Resume Next
For oo = 0 To 8
    txtJ(oo).Text = ""
    dtgCGj.Col = oo + 2
    txtJ(oo).Text = dtgCGj.Text
Next
dtgCGj.Col = 13
lblLId.Caption = dtgCGj.Text
dtgCGj.Col = 14
lblTDid.Caption = dtgCGj.Text
If lblTDid.Caption = "" Then Exit Sub

'绑定替代品资料

Set adoTd = New ADODB.Recordset

tt = "select * from xunjiaTD where tdid=" & Val(lblTDid.Caption)
adoTd.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
txtJ(21).Text = adoTd.Fields("jzPb").Value '机组品牌
txtJ(20).Text = adoTd.Fields("jzxh").Value '机组型号
txtJ(19).Text = adoTd.Fields("yXh").Value '压缩机型号
txtJ(18).Text = adoTd.Fields("CCbh").Value '出厂编号
txtJ(17).Text = adoTd.Fields("jzBh").Value '机组序列号
txtJ(16).Text = adoTd.Fields("ljBh").Value '零件编号
txtJ(15).Text = adoTd.Fields("ljMc").Value '零件名称
txtJ(14).Text = adoTd.Fields("pbcd").Value '品牌产地
txtJ(9).Text = adoTd.Fields("Tdj").Value '替代价格
txtJ(10).Text = adoTd.Fields("tdRen").Value '提供者
txtJ(10).ToolTipText = adoTd.Fields("uid").Value '提供者工号
lblLc.Caption = adoTd.Fields("lc").Value
lblFwid.Caption = adoTd.Fields("Fwid").Value
lblLcRen.Caption = adoTd.Fields("LcRen").Value
lblLcUid.Caption = adoTd.Fields("LcUid").Value
lblGyid.Caption = adoTd.Fields("gyid").Value
If txtJ(9).Text = "" Then
    MsgBox "网络出错,请再试一次,或与马晓聪联系!"
    Call TDQing
    Exit Sub
End If
'绑定供应商资料
adoTd.Close
tt = "select * from xunGy where gyid=" & Val(lblGyid.Caption)
adoTd.Close
adoTd.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
lblGyid.Caption = adoTd.Fields("Gyid").Value
txtGmc.Text = adoTd.Fields("Gmc").Value
txtGAdr.Text = adoTd.Fields("GAdr").Value
txtLcRen1.Text = adoTd.Fields("LcRen1").Value
txtLcRen2.Text = adoTd.Fields("LcRen2").Value
txtT1.Text = adoTd.Fields("T1").Value
txtT2.Text = adoTd.Fields("T2").Value
txtBz.Text = adoTd.Fields("Bz").Value


If lblGyid.Caption = "" Then
    MsgBox "网络出错,请再试一次,或与马晓聪联系!"
    Call TDQing
    Exit Sub
End If
'绑定按钮

tt = "qmrzOpen(58,'" & lblTDid.Caption & "')"
adoTd.Close
adoTd.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc

For oo = 0 To 1

  cmdQm(oo).Tag = adoTd.Fields("zid").Value
  If adoTd.Fields("xf").Value = True Then
       cmdQm(oo).Caption = adoTd.Fields("Qren").Value
       lblTm(oo).Caption = adoTd.Fields("QRQ").Value

  End If
  adoTd.MoveNext
Next

Set adoTd = Nothing
cmdSave.Visible = False
frmMod.Visible = False
frmGy.Visible = False
If lblLcUid.Caption = mod1.DHid Then
    frmMod.Visible = True
    frmGy.Visible = True
End If
If mod1.DName = "张寅" Or mod1.DName = "宋晓炯" Or mod1.DName = "张春华" Then
    frmMod.Visible = True
    frmGy.Visible = True
End If
If mod1.DName = "冯建川" Then
cmdSave.Visible = True
End If
End Sub
