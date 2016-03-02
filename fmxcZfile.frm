VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fmxcZfile 
   BackColor       =   &H00C0FFC0&
   Caption         =   "执行报表"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查 询"
      Height          =   300
      Left            =   11640
      TabIndex        =   11
      Top             =   8400
      Width           =   855
   End
   Begin VB.TextBox txtZ 
      Height          =   270
      Left            =   9840
      TabIndex        =   10
      Top             =   8400
      Width           =   1575
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "fmxcZfile.frx":0000
      Left            =   8280
      List            =   "fmxcZfile.frx":000D
      TabIndex        =   9
      Text            =   "项目名称"
      Top             =   8400
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "导出视图"
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Frame frmEdit 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "frmEdit"
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   8280
      Width           =   5895
      Begin VB.CommandButton cmdDao 
         Caption         =   "导入报表"
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "新建报表"
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除报表"
         Height          =   495
         Left            =   4080
         TabIndex        =   5
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgN 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComDlg.CommonDialog cmdDia 
      Left            =   5280
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   645
      Left            =   12960
      Picture         =   "fmxcZfile.frx":002D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8400
      Width           =   585
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   14208
      _Version        =   393216
      BackColor       =   16777152
      FixedCols       =   0
      BackColorFixed  =   15728356
      BackColorBkg    =   16777152
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
      PictureType     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSAdodcLib.Adodc adoFile 
      Height          =   375
      Left            =   0
      Top             =   8280
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\HMXP9000\work.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\HMXP9000\work.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "worker"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataField       =   "UserId"
      DataSource      =   "adoFile"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   8400
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "fmxcZfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fid As Long
Public Sub dtgbrFF()
dtgBr.Clear
dtgBr.Cols = 19
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "项目名称": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "合同编号": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "合同金额": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "开单金额": dtgBr.CellFontBold = True
dtgBr.Col = 4: dtgBr.Text = "开票金额": dtgBr.CellFontBold = True
dtgBr.Col = 5: dtgBr.Text = "收款金额": dtgBr.CellFontBold = True
dtgBr.Col = 6: dtgBr.Text = "设备收款": dtgBr.CellFontBold = True
dtgBr.Col = 7: dtgBr.Text = "人工收款": dtgBr.CellFontBold = True
dtgBr.Col = 8: dtgBr.Text = "设备收款比例": dtgBr.CellFontBold = True
dtgBr.Col = 9: dtgBr.Text = "人工收款比例": dtgBr.CellFontBold = True
dtgBr.Col = 10: dtgBr.Text = "采购金额": dtgBr.CellFontBold = True
dtgBr.Col = 11: dtgBr.Text = "设备付款": dtgBr.CellFontBold = True
dtgBr.Col = 12: dtgBr.Text = "人工付款": dtgBr.CellFontBold = True
dtgBr.Col = 13: dtgBr.Text = "付款金额": dtgBr.CellFontBold = True
dtgBr.Col = 14: dtgBr.Text = "未付款金额": dtgBr.CellFontBold = True
dtgBr.Col = 15: dtgBr.Text = "人工付款比例": dtgBr.CellFontBold = True
dtgBr.Col = 16: dtgBr.Text = "设备付款比例": dtgBr.CellFontBold = True
dtgBr.Col = 17: dtgBr.Text = "现金流": dtgBr.CellFontBold = True
dtgBr.Col = 18: dtgBr.Text = "Fid": dtgBr.CellFontBold = True

dtgBr.ColWidth(0) = 2500
dtgBr.ColWidth(1) = 1500
dtgBr.RowHeight(0) = dtgBr.RowHeight(1) * 2
dtgN.Clear
dtgN.Cols = 19
dtgBr.ColWidth(18) = 0

End Sub

Private Sub cmdBack_Click()
Me.Visible = False

End Sub

Private Sub cmdDao_Click()
Call Me.InputF(Fid)
Call Me.Bound
Fid = 0
End Sub

Private Sub cmdGx_Click()

End Sub

Private Sub cmdDel_Click()
Dim tt As String
Dim ii As Integer
If Fid = 0 Then Exit Sub
ii = MsgBox("请确认是否删除此报表文件?", vbYesNo + vbQuestion, "请确认!")
If ii = vbNo Then Exit Sub
tt = "update htzfile set delf=0 where fid=" & Fid
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workHT, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = Nothing
Call Me.Bound
Fid = 0

End Sub

Private Sub cmdExport_Click()
Dim ii As Integer
Dim oo As Long
Dim bt() As Byte
Dim tt As String
On Error GoTo BAoM1
If Dir("c:\项目总表.xlsx", vbNormal) <> "" Then
    Kill "c:\项目总表.xlsx"
End If
tt = "select fnr,fsize,fname from htZfile where fname='项目总表.xlsx'"
frmGGL.adoFile.Recordset.Close
frmGGL.adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockReadOnly, adCmdText

ReDim bt(frmGGL.adoFile.Recordset.Fields("Fsize").Value) As Byte
bt() = frmGGL.adoFile.Recordset.Fields("FNR").GetChunk(frmGGL.adoFile.Recordset.Fields("Fsize").Value + 1)

Open ("c:\项目总表.xlsx") For Binary As #2
Put #2, , bt()
Close #2


''''    frmGGL.OLE2.SourceDoc = "c:\项目总表.xlsx"
''''    frmGGL.OLE2.Action = 1
''''    frmGGL.OLE2.DoVerb (-2)
'打开excel，并填充数据
Dim D_Ex As Object
Dim D_ExBook As Object
Dim D_ExSheet As Object

Set D_Ex = CreateObject("Excel.Application")
Set D_ExBook = D_Ex.Workbooks.Open("c:\项目总表.xlsx")    'FullName 是你excel的地址及文件名，如"C：\1.xls"。
Set D_ExSheet = D_ExBook.Worksheets("项目总表")


D_Ex.Visible = False                               'true也行，false看不见excel

On Error Resume Next
For oo = 2 To 5000
    dtgN.Row = oo - 1
    dtgN.Col = 0
    If dtgN.Text = "" Then
        Exit For
    End If

    For ii = 1 To 16

        dtgN.Col = ii - 1

        D_Ex.cells(oo, ii) = dtgN.Text
    Next
Next
D_Ex.Visible = True
D_ExBook.Save '保存
Exit Sub
BAoM1:
MsgBox "出错，请关掉已经打开的excel文件，再试一次！"
On Error Resume Next
D_ExBook.Close '关闭
D_Ex.Quit
End Sub

Private Sub cmdNew_Click()
Dim ii As Integer

Dim bt() As Byte
Dim tt As String
On Error GoTo BAoM1
If Dir("c:\执行报表.xlsm", vbNormal) <> "" Then
    Kill "c:\执行报表.xlsm"
End If
tt = "select fnr,fsize,fname from htZfile where fname='新项目模板.xlsm'"
frmGGL.adoFile.Recordset.Close
frmGGL.adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockReadOnly, adCmdText

ReDim bt(frmGGL.adoFile.Recordset.Fields("Fsize").Value) As Byte
bt() = frmGGL.adoFile.Recordset.Fields("FNR").GetChunk(frmGGL.adoFile.Recordset.Fields("Fsize").Value + 1)

Open ("c:\执行报表.xlsm") For Binary As #2
Put #2, , bt()
Close #2


    frmGGL.OLE2.SourceDoc = "c:\执行报表.xlsm"
    frmGGL.OLE2.Action = 1
    frmGGL.OLE2.DoVerb (-2)
Exit Sub
BAoM1:
MsgBox "出错，请关掉已经打开的excel文件，再试一次！"
On Error Resume Next
'''''D_ExBook.Close '关闭
'''''D_Ex.Quit
End Sub

Private Sub cmdSearch_Click()
Dim tt As String
Dim Ra
Dim La As Long
Dim oo As Long
Dim ii As Integer
Select Case comLx.Text
Case "项目名称"
    tt = "select 项目名称,合同编号,合同金额,开单金额,收款金额,设备收款,人工收款,设备收款比例,人工收款比例," & _
        "采购金额,设备付款,人工付款,付款金额,未付款金额,人工付款比例,设备付款比例,现金流,fid" & _
        " from htZfile where uid='" & mod1.DHid & "' and 项目名称 like '%" & txtZ.Text & "%' and delf=1  order by fid desc"
Case "合同编号"
    tt = "select 项目名称,合同编号,合同金额,开单金额,收款金额,设备收款,人工收款,设备收款比例,人工收款比例," & _
        "采购金额,设备付款,人工付款,付款金额,未付款金额,人工付款比例,设备付款比例,现金流,fid" & _
        " from htZfile where uid='" & mod1.DHid & "' and htbh='" & txtZ.Text & "' and delf=1 order by fid desc"
Case "业务员"
'''    tt = "select 项目名称,合同编号,合同金额,开单金额,收款金额,设备收款,人工收款,设备收款比例,人工收款比例," & _
'''        "采购金额,设备付款,人工付款,付款金额,未付款金额,人工付款比例,设备付款比例,现金流,fid" & _
'''        " from htZfile where uid='" & mod1.DHid & "' and delf=1 order by fid desc"
End Select
If mod1.DName = "顾珣" Or mod1.DName = "乔继敏" Then
Select Case comLx.Text
    Case "项目名称"
        tt = "select 项目名称,合同编号,合同金额,开单金额,收款金额,设备收款,人工收款,设备收款比例,人工收款比例," & _
            "采购金额,设备付款,人工付款,付款金额,未付款金额,人工付款比例,设备付款比例,现金流,fid" & _
            " from htZfile where  项目名称 like '%" & txtZ.Text & "%' and delf=1  order by fid desc"
    Case "合同编号"
        tt = "select 项目名称,合同编号,合同金额,开单金额,收款金额,设备收款,人工收款,设备收款比例,人工收款比例," & _
            "采购金额,设备付款,人工付款,付款金额,未付款金额,人工付款比例,设备付款比例,现金流,fid" & _
            " from htZfile where  htbh='" & txtZ.Text & "' and delf=1 order by fid desc"
    Case "业务员"
    '''    tt = "select 项目名称,合同编号,合同金额,开单金额,收款金额,设备收款,人工收款,设备收款比例,人工收款比例," & _
    '''        "采购金额,设备付款,人工付款,付款金额,未付款金额,人工付款比例,设备付款比例,现金流,fid" & _
    '''        " from htZfile where uid='" & mod1.DHid & "' and delf=1 order by fid desc"
    Case Else
        tt = "select 项目名称,合同编号,合同金额,开单金额,收款金额,设备收款,人工收款,设备收款比例,人工收款比例," & _
            "采购金额,设备付款,人工付款,付款金额,未付款金额,人工付款比例,设备付款比例,现金流,fid" & _
            " from htZfile where delf=1  order by fid desc"
    End Select
    
End If
Call Me.dtgbrFF

Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workHT, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgBr.Rows = La + 30
dtgN.Rows = La + 30

For oo = 1 To La
    dtgBr.Row = oo: dtgN.Row = oo
    For ii = 0 To 17
        dtgBr.Col = ii: dtgBr.Text = Ra(ii, oo - 1)
        dtgN.Col = ii: dtgN.Text = Ra(ii, oo - 1)
    Next
Next
End Sub

Private Sub dtgBr_DblClick()
Dim ii As Integer

Dim bt() As Byte
Dim tt As String
On Error GoTo BaoM2

dtgN.Row = dtgBr.Row
dtgN.Col = 18
Fid = Val(dtgN.Text)
If Fid = 0 Then Exit Sub
If Dir("c:\执行报表.xlsm", vbNormal) <> "" Then
    Kill "c:\执行报表.xlsm"
End If
tt = "select fnr,fsize,fname from htZfile where fid=" & Fid
frmGGL.adoFile.Recordset.Close
frmGGL.adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockReadOnly, adCmdText

ReDim bt(frmGGL.adoFile.Recordset.Fields("Fsize").Value) As Byte
bt() = frmGGL.adoFile.Recordset.Fields("FNR").GetChunk(frmGGL.adoFile.Recordset.Fields("Fsize").Value + 1)

Open ("c:\" & frmGGL.adoFile.Recordset.Fields("fname").Value) For Binary As #2
Put #2, , bt()
Close #2


    frmGGL.OLE2.SourceDoc = "c:\" & frmGGL.adoFile.Recordset.Fields("fname").Value
    frmGGL.OLE2.Action = 1
    frmGGL.OLE2.DoVerb (-2)
    
Exit Sub
BaoM2:
MsgBox "出错，请关掉已经打开的excel文件，再试一次！"
On Error Resume Next
''D_ExBook.Close '关闭
''D_Ex.Quit
End Sub


Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Top = 0: Me.Left = 0
If mod1.DName = "朱婷婷" Or mod1.DName = "王绣霞" Or mod1.DName = "马晓聪" Then
    frmEdit.Visible = True
Else
    frmEdit.Visible = False
End If
End Sub


Public Sub InputF(Fid As Long)
Dim tt As String
Dim bt() As Byte

Dim oo As Integer
Dim FLX As String
Dim Fname As String '文件名(去路径)
Dim 项目名称 As String
Dim 合同编号 As String
Dim 合同金额 As Single
Dim 开单金额 As Single
Dim 开票金额 As Single
Dim 收款金额 As Single
Dim 设备收款 As Single
Dim 人工收款 As Single
Dim 设备收款比例 As Single
Dim 人工收款比例 As Single
Dim 采购金额 As Single
Dim 设备付款 As Single
Dim 人工付款 As Single
Dim 付款金额 As Single
Dim 未付款金额 As Single
Dim 人工付款比例 As Single
Dim 设备付款比例 As Single
Dim 现金流 As Single

On Error GoTo DER77
If mod1.DName = "马晓聪" Then
    cmdDia.ShowOpen
Else
    cmdDia.FileName = "c:\执行报表.xlsm"
End If
'获取文件参数
Dim i As Long

On Error Resume Next
Dim D_Ex As Object
Dim D_ExBook As Object
Dim D_ExSheet As Object

Set D_Ex = CreateObject("Excel.Application")
Set D_ExBook = D_Ex.Workbooks.Open("" & cmdDia.FileName & "")    'FullName 是你excel的地址及文件名，如"C：\1.xls"。
Set D_ExSheet = D_ExBook.Worksheets(0)

D_Ex.Visible = False                               'true也行，false看不见excel

项目名称 = D_Ex.cells(2, 2)
合同编号 = D_Ex.cells(1, 2)
合同金额 = Val(D_Ex.cells(6, 1))
开单金额 = Val(D_Ex.cells(6, 3))
开票金额 = Val(D_Ex.cells(6, 6))
收款金额 = Val(D_Ex.cells(6, 10))
设备收款 = Val(D_Ex.cells(6, 8))
人工收款 = Val(D_Ex.cells(6, 9))
设备收款比例 = Val(D_Ex.cells(6, 12))
人工收款比例 = Val(D_Ex.cells(6, 13))
采购金额 = Val(D_Ex.cells(6, 17))
设备付款 = Val(D_Ex.cells(6, 19))
人工付款 = Val(D_Ex.cells(6, 20))
付款金额 = Val(D_Ex.cells(6, 19)) + Val(D_Ex.cells(6, 20))
未付款金额 = Val(D_Ex.cells(6, 26))
人工付款比例 = Val(D_Ex.cells(6, 29))
设备付款比例 = Val(D_Ex.cells(6, 28))
现金流 = Val(D_Ex.cells(6, 10)) - (Val(D_Ex.cells(6, 19)) + Val(D_Ex.cells(6, 20)))
'D_ExBook.Save '保存
D_ExBook.Close '关闭
D_Ex.Quit

Open cmdDia.FileName For Binary As #1

Fname = ""

For oo = Len(cmdDia.FileName) - 1 To 1 Step -1
    If Mid(cmdDia.FileName, oo, 1) = "\" Then
        Fname = Mid(cmdDia.FileName, oo + 1, Len(cmdDia.FileName) - oo)
        Exit For
        
    End If
Next
'If Right(Fname, 4) = ".xls" Then
'    FLX = Right(Fname, 3)
If Right(Fname, 5) = ".xlsm" Then
    FLX = Right(Fname, 4)
ElseIf Right(Fname, 5) = ".xlsx" Then
    FLX = Right(Fname, 4)
Else
    MsgBox "选择了不正确的文件类型!"
    Exit Sub
End If

On Error Resume Next
ReDim bt(LOF(1) - 1)
'ReDim bt(10000000)
    Get #1, , bt()
If Fid > 0 Then  '更新
    tt = "select * from htZfile where fid=" & Fid
    adoFile.Recordset.Close
    adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockBatchOptimistic, adCmdText
    adoFile.Recordset.Update "Fsize", LOF(1) - 1
    adoFile.Recordset.Update "frq", mod1.DQda
    adoFile.Recordset.Update "项目名称", 项目名称
    adoFile.Recordset.Update "合同编号", 合同编号
    adoFile.Recordset.Update "合同金额", 合同金额
    adoFile.Recordset.Update "开单金额", 开单金额
    adoFile.Recordset.Update "开票金额", 开票金额
    adoFile.Recordset.Update "收款金额", 收款金额
    adoFile.Recordset.Update "设备收款", 设备收款
    adoFile.Recordset.Update "人工收款", 人工收款
    adoFile.Recordset.Update "设备收款比例", 设备收款比例
    adoFile.Recordset.Update "人工收款比例", 人工收款比例
    adoFile.Recordset.Update "采购金额", 采购金额
    adoFile.Recordset.Update "设备付款", 设备付款
    adoFile.Recordset.Update "人工付款", 人工付款
    adoFile.Recordset.Update "付款金额", 付款金额
    adoFile.Recordset.Update "未付款金额", 未付款金额
    adoFile.Recordset.Update "人工付款比例", 人工付款比例
    adoFile.Recordset.Update "设备付款比例", 设备付款比例
    adoFile.Recordset.Update "现金流", 现金流
    adoFile.Recordset.Update "Fname", Fname
    adoFile.Recordset.Fields("FNR").AppendChunk bt()
    adoFile.Recordset.UpdateBatch
    Fid = adoFile.Recordset.Fields("fid").Value
    adoFile.Recordset.Close
    If Fid = 0 Then
        MsgBox "网络故障!"
        Exit Sub
    End If

Else
    tt = "select * from htZfile where fid=0" '添加
    adoFile.Recordset.Close
    adoFile.Recordset.Open tt, mod1.workHT, adOpenKeyset, adLockBatchOptimistic, adCmdText
    adoFile.Recordset.AddNew "ywy", mod1.DName
    adoFile.Recordset.Update "uid", mod1.DHid
    adoFile.Recordset.Update "Fsize", LOF(1) - 1
    adoFile.Recordset.Update "frq", mod1.DQda
    adoFile.Recordset.Update "Fname", Fname
    adoFile.Recordset.Update "项目名称", 项目名称
    adoFile.Recordset.Update "合同编号", 合同编号
    adoFile.Recordset.Update "合同金额", 合同金额
    adoFile.Recordset.Update "开单金额", 开单金额
    adoFile.Recordset.Update "开票金额", 开票金额
    adoFile.Recordset.Update "收款金额", 收款金额
    adoFile.Recordset.Update "设备收款", 设备收款
    adoFile.Recordset.Update "人工收款", 人工收款
    adoFile.Recordset.Update "设备收款比例", 设备收款比例
    adoFile.Recordset.Update "人工收款比例", 人工收款比例
    adoFile.Recordset.Update "采购金额", 采购金额
    adoFile.Recordset.Update "设备付款", 设备付款
    adoFile.Recordset.Update "人工付款", 人工付款
    adoFile.Recordset.Update "付款金额", 付款金额
    adoFile.Recordset.Update "未付款金额", 未付款金额
    adoFile.Recordset.Update "人工付款比例", 人工付款比例
    adoFile.Recordset.Update "设备付款比例", 设备付款比例
    adoFile.Recordset.Update "现金流", 现金流
    adoFile.Recordset.Fields("FNR").AppendChunk bt()
    adoFile.Recordset.UpdateBatch
    Fid = adoFile.Recordset.Fields("fid").Value
    adoFile.Recordset.Close
    If Fid = 0 Then
        MsgBox "网络故障!"
        Exit Sub
    End If


End If
Close #1
MsgBox "成功导入!"

Exit Sub
DER77:
Close #1
End Sub

Public Sub Bound()
Dim tt As String
Dim Ra
Dim La As Long
Dim oo As Long
Dim ii As Integer
dtgBr.Visible = False
Call Me.dtgbrFF
tt = "select 项目名称,合同编号,合同金额,开单金额,开票金额,收款金额,设备收款,人工收款,设备收款比例,人工收款比例," & _
    "采购金额,设备付款,人工付款,付款金额,未付款金额,人工付款比例,设备付款比例,现金流,fid" & _
    " from htZfile where uid='" & mod1.DHid & "' and delf=1 order by fid desc"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workHT, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1

dtgBr.Rows = La + 30
dtgN.Rows = La + 30

For oo = 1 To La
    dtgBr.Row = oo: dtgN.Row = oo
    For ii = 0 To 18
        dtgBr.Col = ii: dtgBr.Text = Ra(ii, oo - 1)
        dtgN.Col = ii: dtgN.Text = Ra(ii, oo - 1)
    Next
Next
dtgBr.Visible = True
End Sub
