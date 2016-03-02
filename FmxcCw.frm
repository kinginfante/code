VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FmxcCw 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFC0&
   Caption         =   "财务评定"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN3 
      Height          =   345
      Left            =   7290
      TabIndex        =   51
      Top             =   2190
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   609
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN2 
      Height          =   315
      Left            =   5430
      TabIndex        =   50
      Top             =   2160
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN1 
      Height          =   315
      Left            =   3690
      TabIndex        =   49
      Top             =   2100
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer timWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   420
   End
   Begin VB.Timer timQuit 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   420
      Top             =   0
   End
   Begin VB.TextBox txtAmount 
      BackColor       =   &H00C0FFFF&
      Height          =   2265
      Left            =   10950
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Text            =   "FmxcCw.frx":0000
      Top             =   150
      Width           =   4215
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   585
      Left            =   14550
      Picture         =   "FmxcCw.frx":0006
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8520
      Width           =   585
   End
   Begin VB.Frame frmAmount 
      BackColor       =   &H00FFFFC0&
      Caption         =   "收款"
      Height          =   6585
      Left            =   10980
      TabIndex        =   5
      Top             =   2580
      Width           =   4245
      Begin VB.CheckBox chkHistory3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "历史记录"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   4860
         Width           =   855
      End
      Begin VB.CommandButton cmdGx3 
         BackColor       =   &H00FF8080&
         Caption         =   "更新"
         Height          =   285
         Left            =   2670
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   5940
         Width           =   735
      End
      Begin VB.CommandButton cmdDel3 
         BackColor       =   &H008080FF&
         Caption         =   "删除"
         Height          =   285
         Left            =   2670
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   6240
         Width           =   735
      End
      Begin VB.TextBox txt3 
         Height          =   270
         Left            =   990
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   6150
         Width           =   1485
      End
      Begin VB.CommandButton cmdAdd3 
         BackColor       =   &H0080C0FF&
         Caption         =   "添加"
         Height          =   315
         Left            =   2670
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   5610
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgRev 
         Height          =   4575
         Left            =   0
         TabIndex        =   19
         Top             =   210
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   8070
         _Version        =   393216
         BackColor       =   16777152
         Rows            =   50
         Cols            =   1
         FixedCols       =   0
         BackColorFixed  =   15728356
         BackColorBkg    =   16777152
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         PictureType     =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComCtl2.DTPicker dtp3 
         Height          =   285
         Left            =   990
         TabIndex        =   45
         Top             =   5820
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         CalendarBackColor=   8454016
         CalendarTitleBackColor=   16711808
         CalendarTrailingForeColor=   -2147483635
         Format          =   109182977
         CurrentDate     =   38797
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "收款时间"
         Height          =   195
         Left            =   120
         TabIndex        =   42
         Top             =   5880
         Width           =   825
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "收款金额"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   6210
         Width           =   855
      End
      Begin VB.Label lblCount3 
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   4860
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "合计"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   22
         Top             =   4860
         Width           =   495
      End
   End
   Begin VB.Frame frmKD 
      BackColor       =   &H00FFFFC0&
      Caption         =   "开单"
      Height          =   6585
      Left            =   7020
      TabIndex        =   4
      Top             =   2580
      Width           =   3975
      Begin VB.CheckBox chkHistory2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "历史记录"
         Height          =   255
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   4860
         Width           =   855
      End
      Begin VB.CommandButton cmdGx2 
         BackColor       =   &H00FF8080&
         Caption         =   "更新"
         Height          =   285
         Left            =   2910
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   5910
         Width           =   795
      End
      Begin VB.CommandButton cmdDel2 
         BackColor       =   &H008080FF&
         Caption         =   "删除"
         Height          =   285
         Left            =   2910
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   6210
         Width           =   795
      End
      Begin VB.TextBox txt2 
         Height          =   270
         Left            =   1050
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   6120
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdd2 
         BackColor       =   &H0080C0FF&
         Caption         =   "添加"
         Height          =   315
         Left            =   2910
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   5580
         Width           =   795
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMD 
         Height          =   4575
         Left            =   0
         TabIndex        =   18
         Top             =   210
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   8070
         _Version        =   393216
         BackColor       =   16777152
         Rows            =   50
         Cols            =   1
         FixedCols       =   0
         BackColorFixed  =   15728356
         BackColorBkg    =   16777152
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         PictureType     =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComCtl2.DTPicker dtp2 
         Height          =   285
         Left            =   1050
         TabIndex        =   44
         Top             =   5760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         CalendarBackColor=   8454016
         CalendarTitleBackColor=   16711808
         CalendarTrailingForeColor=   -2147483635
         Format          =   109248513
         CurrentDate     =   38797
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "开单时间"
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   5820
         Width           =   825
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "开单金额"
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   6180
         Width           =   855
      End
      Begin VB.Label lblCount2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "合计"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   4920
         Width           =   495
      End
   End
   Begin VB.Frame frmkp 
      BackColor       =   &H00FFFFC0&
      Caption         =   "开票"
      Height          =   6585
      Left            =   3030
      TabIndex        =   3
      Top             =   2580
      Width           =   3975
      Begin VB.CheckBox chkHistory1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "历史记录"
         Height          =   255
         Left            =   3030
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4890
         Width           =   855
      End
      Begin VB.CommandButton cmdGx1 
         BackColor       =   &H00FF8080&
         Caption         =   "更新"
         Height          =   285
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   5880
         Width           =   795
      End
      Begin VB.CommandButton cmdDel1 
         BackColor       =   &H008080FF&
         Caption         =   "删除"
         Height          =   285
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   6180
         Width           =   795
      End
      Begin VB.TextBox txtFP 
         Height          =   270
         Left            =   1140
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   6120
         Width           =   1665
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Left            =   1140
         TabIndex        =   29
         Text            =   "Text1"
         Top             =   5790
         Width           =   1665
      End
      Begin VB.CommandButton cmdAdd1 
         BackColor       =   &H0080C0FF&
         Caption         =   "添加"
         Height          =   315
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5550
         Width           =   795
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgMI 
         Height          =   4575
         Left            =   30
         TabIndex        =   17
         Top             =   210
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   8070
         _Version        =   393216
         BackColor       =   16777152
         Rows            =   50
         Cols            =   1
         FixedCols       =   0
         BackColorFixed  =   15728356
         BackColorBkg    =   16777152
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   1
         PictureType     =   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComCtl2.DTPicker dtp1 
         Height          =   285
         Left            =   1140
         TabIndex        =   43
         Top             =   5460
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         CalendarBackColor=   8454016
         CalendarTitleBackColor=   16711808
         CalendarTrailingForeColor=   -2147483635
         Format          =   109248513
         CurrentDate     =   38797
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "开票时间"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   5490
         Width           =   825
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "发票号码"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   6180
         Width           =   795
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "开票金额"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   5850
         Width           =   855
      End
      Begin VB.Label lblCount1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label5"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   810
         TabIndex        =   23
         Top             =   4890
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "合计"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Top             =   4920
         Width           =   495
      End
   End
   Begin VB.TextBox txtHtbh 
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   180
      TabIndex        =   1
      Top             =   660
      Width           =   2325
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgHt 
      Height          =   8145
      Left            =   180
      TabIndex        =   2
      Top             =   1020
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   14367
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   50
      Cols            =   1
      FixedCols       =   0
      BackColorFixed  =   15728356
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   1
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Line Line1 
      X1              =   2850
      X2              =   2850
      Y1              =   9120
      Y2              =   0
   End
   Begin VB.Label lblSales 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label10"
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   1470
      Width           =   3105
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "业务员"
      Height          =   225
      Left            =   3240
      TabIndex        =   15
      Top             =   1530
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "应收款项"
      Height          =   195
      Left            =   9990
      TabIndex        =   13
      Top             =   210
      Width           =   915
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label7"
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   1035
      Width           =   3105
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "合同金额"
      Height          =   255
      Left            =   3180
      TabIndex        =   11
      Top             =   1095
      Width           =   885
   End
   Begin VB.Label lblSDbh 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      Height          =   255
      Left            =   4080
      TabIndex        =   10
      Top             =   615
      Width           =   3105
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "速达编号"
      Height          =   225
      Left            =   3180
      TabIndex        =   9
      Top             =   640
      Width           =   1155
   End
   Begin VB.Label lblCustomer 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   180
      Width           =   3105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "客户名称"
      Height          =   225
      Left            =   3180
      TabIndex        =   7
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "合同编号"
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   360
      Width           =   945
   End
End
Attribute VB_Name = "FmxcCw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Hid As Long
Dim timZm As Integer '数据提交后,由timWait执行的后续命令ID(1 开票编辑,2开单编辑,3收款编辑),

Dim Id As Long

Private Sub chkHistory1_Click()
Dim tt As String
Dim Ra
Dim La As Integer
If chkHistory1.Value = 0 Then
    tt = "select rq,amount,hm,id from htpingKd where hid=" & Me.Hid & " and lb='开票' order by Id desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Else
    tt = "SELECT dbo.MLMX.mt2, dbo.MLMX.mD1, dbo.MLMX.mM1, dbo.MLMX.mt3,HMText.dbo.Ml.trq,dbo.MLMX.md2,dbo.MLMX.mm3,dbo.MLMX.mt4 FROM dbo.MLMX INNER JOIN " & _
        "HMText.dbo.ML ON dbo.MLMX.zid = HMText.dbo.ML.Zid WHERE (HMText.dbo.ML.NB = '财务评定') and HMText.dbo.ML.NBLX='开票编辑' and HMText.dbo.ML.bh='" & Me.Hid & "' and HMtext.dbo.ML.cf=1 order by HMText.dbo.ML.zid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.wzcc, adOpenForwardOnly, adLockReadOnly, adCmdText
End If
        On Error Resume Next
        Ra = mod1.HTP.GetRows
        La = UBound(Ra, 2) + 1
        Call Me.MiBound(Ra, La)
        mod1.HTP.Close
        Set mod1.HTP = Nothing
End Sub

Private Sub chkHistory2_Click()
Dim tt As String
Dim Ra
Dim La As Integer
If chkHistory2.Value = 0 Then
    tt = "select rq,amount,hm,id from htpingKd where hid=" & Me.Hid & " and lb='开单' order by Id desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Else
    tt = "SELECT dbo.MLMX.mt2, dbo.MLMX.mD1, dbo.MLMX.mM1, dbo.MLMX.mt3,HMText.dbo.Ml.trq,dbo.MLMX.md2,dbo.MLMX.mm3,dbo.MLMX.mt4 FROM dbo.MLMX INNER JOIN " & _
        "HMText.dbo.ML ON dbo.MLMX.zid = HMText.dbo.ML.Zid WHERE (HMText.dbo.ML.NB = '财务评定') and HMText.dbo.ML.NBLX='开单编辑' and HMText.dbo.ML.bh='" & Me.Hid & "' and HMtext.dbo.ML.cf=1 order by HMText.dbo.ML.zid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.wzcc, adOpenForwardOnly, adLockReadOnly, adCmdText
End If
        On Error Resume Next
        Ra = mod1.HTP.GetRows
        La = UBound(Ra, 2) + 1
        Call Me.MdBound(Ra, La)
        mod1.HTP.Close
        Set mod1.HTP = Nothing
End Sub

Private Sub chkHistory3_Click()
Dim tt As String
Dim Ra
Dim La As Integer
If chkHistory3.Value = 0 Then
    tt = "select rq,amount,hm,id from htpingKd where hid=" & Me.Hid & " and lb='收款' order by Id desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Else
    tt = "SELECT dbo.MLMX.mt2, dbo.MLMX.mD1, dbo.MLMX.mM1, dbo.MLMX.mt3,HMText.dbo.Ml.trq,dbo.MLMX.md2,dbo.MLMX.mm3,dbo.MLMX.mt4 FROM dbo.MLMX INNER JOIN " & _
        "HMText.dbo.ML ON dbo.MLMX.zid = HMText.dbo.ML.Zid WHERE (HMText.dbo.ML.NB = '财务评定') and HMText.dbo.ML.NBLX='收款编辑' and HMText.dbo.ML.bh='" & Me.Hid & "' and HMtext.dbo.ML.cf=1 order by HMText.dbo.ML.zid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.wzcc, adOpenForwardOnly, adLockReadOnly, adCmdText
End If
        On Error Resume Next
        Ra = mod1.HTP.GetRows
        La = UBound(Ra, 2) + 1
        Call Me.RevBound(Ra, La)
        mod1.HTP.Close
        Set mod1.HTP = Nothing
End Sub

Private Sub cmdAdd1_Click()

Dim tt As String


Dim oo As Integer
On Error Resume Next

If Me.chkHistory1.Value = 1 Then Exit Sub
If Val(txt1.Text) = 0 Then
Exit Sub
End If



timZm = 1 '开票编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "财务评定"
    mod1.cmd.Parameters("@NBLX") = "开票编辑"
    mod1.cmd.Parameters("@bh") = Me.Hid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = "添加"
    mod1.cmd.Parameters("@mt3") = txtFP.Text '开票号码
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt1.Text)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = dtp1.Value

    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 3 Then '保存

        End If
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
Me.chkHistory1.Value = 0
End Sub

Private Sub cmdAdd2_Click()
Dim tt As String


Dim oo As Integer
On Error Resume Next
If Me.chkHistory1.Value = 2 Then Exit Sub
If Val(txt2.Text) = 0 Then
Exit Sub
End If



timZm = 2 '开票编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "财务评定"
    mod1.cmd.Parameters("@NBLX") = "开单编辑"
    mod1.cmd.Parameters("@bh") = Me.Hid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = "添加"
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt2.Text)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = dtp2.Value

    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 3 Then '保存

        End If
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
Me.chkHistory2.Value = 0
End Sub


Private Sub cmdAdd3_Click()
Dim tt As String


Dim oo As Integer
On Error Resume Next

If Me.chkHistory1.Value = 3 Then Exit Sub
If Val(txt3.Text) = 0 Then
Exit Sub
End If



timZm = 3 '收款编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "财务评定"
    mod1.cmd.Parameters("@NBLX") = "收款编辑"
    mod1.cmd.Parameters("@bh") = Me.Hid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = "添加"
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt3.Text)
    mod1.cmd.Parameters("@mm2") = 0
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = dtp3.Value

    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 3 Then '保存

        End If
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
Set mod1.cmd = Nothing
Me.chkHistory3.Value = 0
End Sub


Private Sub cmdBack_Click()
Me.Visible = False
frmZu.Enabled = True
End Sub

Private Sub cmdDel1_Click()
Dim tt As String

Dim ii As Integer
Dim oo As Integer
On Error Resume Next

If Me.chkHistory1.Value = 1 Then Exit Sub
If Id = 0 Then
Exit Sub
End If
ii = MsgBox("是否删除此记录?", vbQuestion + vbYesNo, "询问")
If ii = vbNo Then
    Exit Sub
End If


timZm = 1 '开票编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "财务评定"
    mod1.cmd.Parameters("@NBLX") = "开票编辑"
    mod1.cmd.Parameters("@bh") = Me.Hid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = "删除"
    mod1.cmd.Parameters("@mt3") = txtFP.Text '开票号码
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt1.Text)
    mod1.cmd.Parameters("@mm2") = Id
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = dtp1.Value

    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 3 Then '保存

        End If
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
Me.chkHistory1.Value = 0
End Sub

Private Sub cmdDel2_Click()
Dim tt As String

Dim ii As Integer
Dim oo As Integer
On Error Resume Next
If Me.chkHistory1.Value = 2 Then Exit Sub
If Id = 0 Then
Exit Sub
End If
ii = MsgBox("是否删除此记录?", vbQuestion + vbYesNo, "询问")
If ii = vbNo Then
    Exit Sub
End If


timZm = 2 '开单编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "财务评定"
    mod1.cmd.Parameters("@NBLX") = "开单编辑"
    mod1.cmd.Parameters("@bh") = Me.Hid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = "删除"
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt2.Text)
    mod1.cmd.Parameters("@mm2") = Id
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = dtp2.Value

    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 3 Then '保存

        End If
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
Set mod1.cmd = Nothing
Me.chkHistory2.Value = 0
End Sub


Private Sub cmdDel3_Click()
Dim tt As String

Dim ii As Integer
Dim oo As Integer
On Error Resume Next
If Me.chkHistory1.Value = 3 Then Exit Sub
If Id = 0 Then
Exit Sub
End If
ii = MsgBox("是否删除此记录?", vbQuestion + vbYesNo, "询问")
If ii = vbNo Then
    Exit Sub
End If


timZm = 3 '开票编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "财务评定"
    mod1.cmd.Parameters("@NBLX") = "收款编辑"
    mod1.cmd.Parameters("@bh") = Me.Hid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = "删除"
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt3.Text)
    mod1.cmd.Parameters("@mm2") = Id
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = dtp3.Value

    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 3 Then '保存

        End If
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
Me.chkHistory3.Value = 0
End Sub


Private Sub cmdGx1_Click()
Dim tt As String


Dim oo As Integer
On Error Resume Next

If Me.chkHistory1.Value = 1 Then Exit Sub
If Id = 0 Then
Exit Sub
End If



timZm = 1 '开票编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "财务评定"
    mod1.cmd.Parameters("@NBLX") = "开票编辑"
    mod1.cmd.Parameters("@bh") = Me.Hid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = "更新"
    mod1.cmd.Parameters("@mt3") = txtFP.Text '开票号码
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt1.Text)
    mod1.cmd.Parameters("@mm2") = Id
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = dtp1.Value

    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 3 Then '保存

        End If
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
Me.chkHistory1.Value = 0
End Sub

Private Sub cmdGx2_Click()
Dim tt As String


Dim oo As Integer
On Error Resume Next
If Me.chkHistory1.Value = 2 Then Exit Sub
If Id = 0 Then
Exit Sub
End If



timZm = 2 '开票编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "财务评定"
    mod1.cmd.Parameters("@NBLX") = "开单编辑"
    mod1.cmd.Parameters("@bh") = Me.Hid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = "更新"
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt2.Text)
    mod1.cmd.Parameters("@mm2") = Id
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = dtp2.Value

    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 3 Then '保存

        End If
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
Set mod1.cmd = Nothing
Me.chkHistory2.Value = 0
End Sub

Private Sub cmdGx3_Click()
Dim tt As String


Dim oo As Integer
On Error Resume Next
If Me.chkHistory1.Value = 3 Then Exit Sub
If Id = 0 Then
Exit Sub
End If



timZm = 3 '开票编辑
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "MLAdd"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@zid") = 0
    mod1.cmd.Parameters("@errch") = ""
    mod1.cmd.Parameters("@NB") = "财务评定"
    mod1.cmd.Parameters("@NBLX") = "收款编辑"
    mod1.cmd.Parameters("@bh") = Me.Hid
    mod1.cmd.Parameters("@ywy") = mod1.DName
    mod1.cmd.Parameters("@uid") = mod1.DHid
    mod1.cmd.Parameters("@mt1") = ""
    mod1.cmd.Parameters("@mt2") = "更新"
    mod1.cmd.Parameters("@mt3") = ""
    mod1.cmd.Parameters("@mlt1") = ""
    mod1.cmd.Parameters("@mm1") = Val(txt3.Text)
    mod1.cmd.Parameters("@mm2") = Id
    mod1.cmd.Parameters("@mb1") = 0
    mod1.cmd.Parameters("@md1") = dtp3.Value

    Call mod1.REV: mod1.cmd.Parameters("@zt") = mod1.ZT
    mod1.cmd.Execute
    mod1.Zid = mod1.cmd.Parameters("@zid").Value
    If mod1.cmd.Parameters("@errch").Value <> "成功" Then
        MsgBox "网络出现故障,请试着关闭程序,再执行此操作,如果仍然失败,请与马晓聪联系!"
        If timZm = 3 Then '保存

        End If
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
Me.chkHistory3.Value = 0
End Sub


Private Sub dtgHt_DblClick()
If dtgHt.Row <> 0 And dtgHt.Text <> "" Then
    txtHtbh.Text = dtgHt.Text
    txtHtbh.ForeColor = &HC00000
    Call Me.Bound(txtHtbh.Text)
End If
End Sub

Private Sub dtgMD_Click()
If Me.chkHistory2.Value = 1 Then Exit Sub
dtgN2.Row = dtgMD.Row
dtgN2.Col = 3: Id = Val(dtgN2.Text)
If Id = 0 Then Exit Sub
dtgN2.Col = 0: dtp2.Value = dtgN2.Text
dtgN2.Col = 1: txt2.Text = dtgN2.Text

End Sub

Private Sub dtgMI_Click()
If Me.chkHistory1.Value = 1 Then Exit Sub
dtgN1.Row = dtgMI.Row
dtgN1.Col = 3: Id = Val(dtgN1.Text)
If Id = 0 Then Exit Sub
dtgN1.Col = 0: dtp1.Value = dtgN1.Text
dtgN1.Col = 1: txt1.Text = dtgN1.Text
dtgN1.Col = 2: txtFP.Text = dtgN1.Text

End Sub

Private Sub dtgRev_Click()
If Me.chkHistory3.Value = 1 Then Exit Sub
dtgN3.Row = dtgRev.Row
dtgN3.Col = 3: Id = Val(dtgN3.Text)
If Id = 0 Then Exit Sub
dtgN3.Col = 0: dtp3.Value = dtgN3.Text
dtgN3.Col = 1: txt3.Text = dtgN3.Text
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
Me.txtHtbh.Text = ""
Call Me.HtInitialize
Call Me.Initialize
dtp1.Value = Date
dtp2.Value = Date
dtp3.Value = Date

End Sub

Public Sub HtInitialize()
dtgHt.Clear
dtgHt.Cols = 1
dtgHt.ColWidth(0) = 2040
dtgHt.Row = 0: dtgHt.Col = 0: dtgHt.Text = "合同编号选择": dtgHt.CellFontBold = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

Me.Visible = False
Cancel = True
frmZu.Enabled = True
End Sub

Private Sub MSHFlexGrid3_Click()

End Sub

Private Sub timQuit_Timer()
Dim oo As Integer
Dim ii As Integer
On Error Resume Next
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
Dim tt As String
If timZm = 1 Then '如果为添加合同评审

End If
timQuit.Enabled = False

End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
Dim oo As Integer
Dim Ra
Dim La As Integer
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
        tt = "select rq,amount,hm,id from htpingKd where hid=" & Me.Hid & " and lb='开票' order by Id desc"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        La = UBound(Ra, 2) + 1
        Call Me.MiBound(Ra, La)
    ElseIf timZm = 2 Then
        tt = "select rq,amount,hm,id from htpingKd where hid=" & Me.Hid & " and lb='开单' order by Id desc"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        La = UBound(Ra, 2) + 1
        Call Me.MdBound(Ra, La)
    ElseIf timZm = 3 Then
        tt = "select rq,amount,hm,id from htpingKd where hid=" & Me.Hid & " and lb='收款' order by Id desc"
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        La = UBound(Ra, 2) + 1
        Call Me.RevBound(Ra, La)
    End If
    mod1.HTP.Close
    Set mod1.HTP = Nothing
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

Private Sub txtHtbh_Change()
Dim Ra
Dim La
Dim tt As String
Dim oo As Integer
If Len(Trim(txtHtbh.Text)) < 3 Then Exit Sub
tt = "select htbh from htping where htbh like '%" & Trim(txtHtbh.Text) & "%'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
Call Me.HtInitialize
dtgHt.Rows = La + 30
dtgHt.Col = 0
For oo = 1 To La
    dtgHt.Row = oo
    dtgHt.Text = Ra(0, oo - 1)
Next
txtHtbh.ForeColor = &H80000008
End Sub

Private Sub txtHtbh_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    dtgHt.Row = 1
    If dtgHt.Text <> "" Then
        txtHtbh.Text = dtgHt.Text
        txtHtbh.ForeColor = &HC00000
        Call Me.Bound(txtHtbh.Text)
    End If
End If
End Sub

Public Sub Initialize()
Me.chkHistory1.Value = False
Me.chkHistory2.Value = False
Me.chkHistory3.Value = False
Me.lblCustomer.Caption = ""
Me.lblSDbh.Caption = ""
Me.lblAmount.Caption = ""
Me.lblSales.Caption = ""
Me.txtAmount.Text = ""
lblCount1.Caption = ""
lblCount2.Caption = ""
lblCount3.Caption = ""
txtFP.Text = ""
txt1.Text = ""
txt2.Text = ""
txt3.Text = ""
Call Me.MiInitialize
Call Me.MDInitialize
Call Me.REVInitialize
End Sub

Public Sub MiInitialize()
Dim oo As Integer
dtgMI.Cols = 4
dtgMI.Clear

If chkHistory1.Value = 1 Then
    dtgMI.Row = 0
    dtgMI.Col = 0: dtgMI.Text = "操作": dtgMI.CellFontBold = True
    dtgMI.Col = 1: dtgMI.Text = "数据": dtgMI.CellFontBold = True
    dtgMI.Col = 2: dtgMI.Text = "更新数据": dtgMI.CellFontBold = True
    dtgMI.ColWidth(3) = 0
    dtgMI.ColWidth(2) = 0
    dtgMI.ColWidth(1) = 2595
    dtgMI.ColWidth(0) = -1
Else
    dtgMI.Row = 0
    dtgMI.Col = 0: dtgMI.Text = "开票时间": dtgMI.CellFontBold = True
    dtgMI.Col = 1: dtgMI.Text = "金额": dtgMI.CellFontBold = True
    dtgMI.Col = 2: dtgMI.Text = "发票号码": dtgMI.CellFontBold = True
    dtgMI.ColWidth(3) = 0
    dtgMI.ColWidth(2) = 1590
    dtgMI.ColWidth(1) = 1050
    dtgMI.ColWidth(0) = -1
    lblCount1.Caption = 0
    dtgN1.Clear
    dtgN1.Cols = 4
    dtgN1.Rows = dtgMI.Rows
    For oo = 1 To dtgMI.Rows - 1
        dtgMI.RowHeight(oo) = dtgMI.RowHeight(0)
    Next
End If
End Sub
Public Sub MDInitialize()
Dim oo As Integer
dtgMD.Cols = 4
dtgMD.Clear
If chkHistory2.Value = 1 Then
    dtgMD.Row = 0
    dtgMD.Col = 0: dtgMD.Text = "操作": dtgMD.CellFontBold = True
    dtgMD.Col = 1: dtgMD.Text = "数据": dtgMD.CellFontBold = True
    dtgMD.Col = 2: dtgMD.Text = "更新数据": dtgMD.CellFontBold = True
    dtgMD.ColWidth(3) = 0
    dtgMD.ColWidth(2) = 0
    dtgMD.ColWidth(1) = 2595
    dtgMD.ColWidth(0) = -1
Else
    dtgMD.Row = 0
    dtgMD.Col = 0: dtgMD.Text = "开单时间": dtgMD.CellFontBold = True
    dtgMD.Col = 1: dtgMD.Text = "金额": dtgMD.CellFontBold = True
    dtgMD.Col = 2: dtgMD.Text = "单号": dtgMD.CellFontBold = True
    dtgMD.ColWidth(3) = 0
    dtgMD.ColWidth(2) = 0
    dtgMD.ColWidth(1) = 1050
    lblCount2.Caption = 0
    dtgN2.Clear
    dtgN2.Cols = 4
    dtgN2.Rows = dtgMD.Rows
    For oo = 1 To dtgMD.Rows - 1
        dtgMD.RowHeight(oo) = dtgMD.RowHeight(0)
    Next
End If
End Sub
Public Sub REVInitialize()
Dim oo As Integer
dtgRev.Cols = 4
dtgRev.Clear
If chkHistory3.Value = 1 Then
    dtgRev.Row = 0
    dtgRev.Col = 0: dtgRev.Text = "操作": dtgRev.CellFontBold = True
    dtgRev.Col = 1: dtgRev.Text = "数据": dtgRev.CellFontBold = True
    dtgRev.Col = 2: dtgRev.Text = "更新数据": dtgRev.CellFontBold = True
    dtgRev.ColWidth(3) = 0
    dtgRev.ColWidth(2) = 0
    dtgRev.ColWidth(1) = 2595
    dtgRev.ColWidth(0) = -1
Else
    dtgRev.Row = 0
    dtgRev.Col = 0: dtgRev.Text = "收款时间": dtgRev.CellFontBold = True
    dtgRev.Col = 1: dtgRev.Text = "金额": dtgRev.CellFontBold = True
    dtgRev.Col = 2: dtgRev.Text = "单号": dtgRev.CellFontBold = True
    dtgRev.ColWidth(3) = 0
    dtgRev.ColWidth(2) = 0
    dtgRev.ColWidth(1) = 1050
    lblCount3.Caption = 0
    dtgN3.Clear
    dtgN3.Cols = 4
    dtgN3.Rows = dtgRev.Rows
    For oo = 1 To dtgRev.Rows - 1
        dtgRev.RowHeight(oo) = dtgRev.RowHeight(0)
    Next
End If
End Sub
Public Sub Bound(Bh As String)
Dim oo As Integer
Dim tt As String
Dim Ra, Rb, RC, RD, RE
Dim Lb As Integer
Dim Lc As Integer
Dim Ld As Integer
Dim Le As Integer
Call Me.Initialize
tt = "Declare @hid int;" & _
    "select @hid=hid from htping where htbh='" & Bh & "';" & _
    "select khmc,zbh,htze,xywy,hid from htping where htbh='" & Bh & "';" & _
    "select rq,yingfje from htping1 where htbh=cast(@hid as nvarchar(20));" & _
    "select rq,amount,hm,id from htpingKd where hid=@hid and lb='开票' order by Id desc;" & _
    "select rq,amount,hm,id from htpingKd where hid=@hid and lb='开单' order by Id desc;" & _
    "select rq,amount,hm,id from htpingKd where hid=@hid and lb='收款' order by Id desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RC = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RD = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
RE = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Lb = UBound(Rb, 2) + 1
Lc = UBound(RC, 2) + 1
Ld = UBound(RD, 2) + 1
Le = UBound(RE, 2) + 1
lblCustomer.Caption = Ra(0, 0)
lblSDbh.Caption = Ra(1, 0)
lblAmount.Caption = Ra(2, 0)
lblSales.Caption = Ra(3, 0)
Me.Hid = Ra(4, 0)
For oo = 0 To Lb
    txtAmount.Text = txtAmount.Text & "(" & Trim(Str(oo + 1)) & ") " & Rb(0, oo) & ": "
    txtAmount.Text = txtAmount.Text & Rb(1, oo) & Chr(13) & Chr(10)
Next
Call Me.MiBound(RC, Lc)
Call Me.MdBound(RD, Ld)
Call Me.RevBound(RE, Le)
End Sub

Public Sub MiBound(Ra, La As Integer)
Dim oo As Integer
On Error Resume Next
dtgMI.Visible = False
Call Me.MiInitialize
If Me.chkHistory1.Value = 1 Then
    For oo = 1 To La
        dtgMI.Row = oo
        dtgMI.Col = 0: dtgMI.Text = Ra(0, oo - 1) & Chr(13) & Chr(10) & Ra(4, oo - 1)
        dtgMI.Col = 1: dtgMI.Text = "日期: " & Ra(1, oo - 1) & Chr(13) & Chr(10) & "金额: " & Ra(2, oo - 1) & Chr(13) & Chr(10) & "发票: " & Ra(3, oo - 1)
        dtgMI.RowHeight(oo) = dtgMI.RowHeight(0) * 3
        If Ra(0, oo - 1) = "更新" Then
            dtgMI.RowHeight(oo) = dtgMI.RowHeight(0) * 4
            dtgMI.Col = 1
            dtgMI.Text = "日期: " & Ra(1, oo - 1) & "(" & Ra(5, oo - 1) & ")" & Chr(13) & Chr(10) & _
                        "金额: " & Ra(2, oo - 1) & "(" & Ra(6, oo - 1) & ")" & Chr(13) & Chr(10) & _
                        "发票: " & Ra(3, oo - 1) & "(" & Ra(7, oo - 1) & ")"
        ElseIf Ra(0, oo - 1) = "删除" Then
            'dtgMI.RowHeight(oo) = dtgMI.RowHeight(0) * 4
            dtgMI.Col = 1
            dtgMI.Text = "日期: " & Ra(5, oo - 1) & Chr(13) & Chr(10) & _
                        "金额: " & Ra(6, oo - 1) & Chr(13) & Chr(10) & _
                        "发票: " & Ra(7, oo - 1)
        End If
    Next
Else
    For oo = 1 To La
        dtgMI.Row = oo
        dtgMI.Col = 0: dtgMI.Text = Ra(0, oo - 1)
        dtgMI.Col = 1: dtgMI.Text = Ra(1, oo - 1)
        lblCount1.Caption = Val(lblCount1.Caption) + Val(dtgMI.Text)
        dtgMI.Col = 2: dtgMI.Text = Ra(2, oo - 1)
        dtgMI.Col = 3: dtgMI.Text = Ra(3, oo - 1)
        dtgN1.Row = oo
        dtgN1.Col = 0: dtgN1.Text = Ra(0, oo - 1)
        dtgN1.Col = 1: dtgN1.Text = Ra(1, oo - 1)
        dtgN1.Col = 2: dtgN1.Text = Ra(2, oo - 1)
        dtgN1.Col = 3: dtgN1.Text = Ra(3, oo - 1)
        dtgMI.RowHeight(oo) = dtgMI.RowHeight(0)
    Next
End If
dtgMI.Visible = True
dtgMI.TopRow = 1
End Sub

Public Sub MdBound(Ra, La As Integer)
Dim oo As Integer
On Error Resume Next
dtgMD.Visible = False
Call Me.MDInitialize
If Me.chkHistory2.Value = 1 Then
    For oo = 1 To La
        dtgMD.Row = oo
        dtgMD.Col = 0: dtgMD.Text = Ra(0, oo - 1) & Chr(13) & Chr(10) & Ra(4, oo - 1)
        'dtgMD.Col = 1: dtgMD.Text = "日期: " & Ra(1, oo - 1) & Chr(13) & Chr(10) & "金额: " & Ra(2, oo - 1) & Chr(13) & Chr(10) & "发票: " & Ra(3, oo - 1)
        dtgMD.Col = 1: dtgMD.Text = "日期: " & Ra(1, oo - 1) & Chr(13) & Chr(10) & "金额: " & Ra(2, oo - 1)
        dtgMD.RowHeight(oo) = dtgMD.RowHeight(0) * 3
        If Ra(0, oo - 1) = "更新" Then
            dtgMD.RowHeight(oo) = dtgMD.RowHeight(0) * 4
            dtgMD.Col = 1
            'dtgMD.Text = "日期: " & Ra(1, oo - 1) & "(" & Ra(5, oo - 1) & ")" & Chr(13) & Chr(10) & _
                        "金额: " & Ra(2, oo - 1) & "(" & Ra(6, oo - 1) & ")" & Chr(13) & Chr(10) & _
                        "发票: " & Ra(3, oo - 1) & "(" & Ra(7, oo - 1) & ")"
            dtgMD.Text = "日期: " & Ra(1, oo - 1) & "(" & Ra(5, oo - 1) & ")" & Chr(13) & Chr(10) & _
                        "金额: " & Ra(2, oo - 1) & "(" & Ra(6, oo - 1) & ")"
        End If
    Next
Else
    For oo = 1 To La + 1
        dtgMD.Row = oo
        dtgMD.Col = 0: dtgMD.Text = Ra(0, oo - 1)
        dtgMD.Col = 1: dtgMD.Text = Ra(1, oo - 1)
        lblCount2.Caption = Val(lblCount2.Caption) + Val(dtgMD.Text)
        dtgMD.Col = 2: dtgMD.Text = Ra(2, oo - 1)
        dtgMD.Col = 3: dtgMD.Text = Ra(3, oo - 1)
        dtgN2.Row = oo
        dtgN2.Col = 0: dtgN2.Text = Ra(0, oo - 1)
        dtgN2.Col = 1: dtgN2.Text = Ra(1, oo - 1)
        dtgN2.Col = 2: dtgN2.Text = Ra(2, oo - 1)
        dtgN2.Col = 3: dtgN2.Text = Ra(3, oo - 1)
    Next
End If
dtgMD.Visible = True
dtgMD.TopRow = 1
End Sub

Public Sub RevBound(Ra, La As Integer)
Dim oo As Integer
On Error Resume Next
dtgRev.Visible = False
Call Me.REVInitialize
If Me.chkHistory3.Value = 1 Then
    For oo = 1 To La
        dtgRev.Row = oo
        dtgRev.Col = 0: dtgRev.Text = Ra(0, oo - 1) & Chr(13) & Chr(10) & Ra(4, oo - 1)
        'dtgrev.Col = 1: dtgrev.Text = "日期: " & Ra(1, oo - 1) & Chr(13) & Chr(10) & "金额: " & Ra(2, oo - 1) & Chr(13) & Chr(10) & "发票: " & Ra(3, oo - 1)
        dtgRev.Col = 1: dtgRev.Text = "日期: " & Ra(1, oo - 1) & Chr(13) & Chr(10) & "金额: " & Ra(2, oo - 1)
        dtgRev.RowHeight(oo) = dtgRev.RowHeight(0) * 3
        If Ra(0, oo - 1) = "更新" Then
            dtgRev.RowHeight(oo) = dtgRev.RowHeight(0) * 4
            dtgRev.Col = 1
            'dtgrev.Text = "日期: " & Ra(1, oo - 1) & "(" & Ra(5, oo - 1) & ")" & Chr(13) & Chr(10) & _
                        "金额: " & Ra(2, oo - 1) & "(" & Ra(6, oo - 1) & ")" & Chr(13) & Chr(10) & _
                        "发票: " & Ra(3, oo - 1) & "(" & Ra(7, oo - 1) & ")"
            dtgRev.Text = "日期: " & Ra(1, oo - 1) & "(" & Ra(5, oo - 1) & ")" & Chr(13) & Chr(10) & _
                        "金额: " & Ra(2, oo - 1) & "(" & Ra(6, oo - 1) & ")"
        End If
    Next
Else
    For oo = 1 To La + 1
        dtgRev.Row = oo
        dtgRev.Col = 0: dtgRev.Text = Ra(0, oo - 1)
        dtgRev.Col = 1: dtgRev.Text = Ra(1, oo - 1)
        lblCount3.Caption = Val(lblCount3.Caption) + Val(dtgRev.Text)
        dtgRev.Col = 2: dtgRev.Text = Ra(2, oo - 1)
        dtgRev.Col = 3: dtgRev.Text = Ra(3, oo - 1)
        dtgN3.Row = oo
        dtgN3.Col = 0: dtgN3.Text = Ra(0, oo - 1)
        dtgN3.Col = 1: dtgN3.Text = Ra(1, oo - 1)
        dtgN3.Col = 2: dtgN3.Text = Ra(2, oo - 1)
        dtgN3.Col = 3: dtgN3.Text = Ra(3, oo - 1)
    Next
End If
dtgRev.Visible = True
dtgRev.TopRow = 1
End Sub
