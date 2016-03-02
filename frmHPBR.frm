VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmHPBR 
   Caption         =   "货品资料查询"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10950
   LinkTopic       =   "Form2"
   ScaleHeight     =   5625
   ScaleWidth      =   10950
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame frmZX 
      Caption         =   "自选成本（人工类成本内容）"
      Height          =   855
      Left            =   3240
      TabIndex        =   28
      Top             =   4320
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.CommandButton cmdGB 
      Caption         =   "关闭"
      Height          =   315
      Left            =   10200
      TabIndex        =   12
      Top             =   5280
      Width           =   735
   End
   Begin VB.Frame frmC 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   4920
      Width           =   9015
      Begin VB.Frame frmDao 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   5040
         TabIndex        =   26
         Top             =   0
         Width           =   3495
         Begin VB.CommandButton Command1 
            Caption         =   "近期导入"
            Height          =   255
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame frmXT 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   345
         Left            =   4200
         TabIndex        =   23
         Top             =   5000
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton cmdXT 
            Caption         =   "查询相同记录"
            Height          =   315
            Left            =   840
            TabIndex        =   25
            Top             =   150
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.ComboBox comMLx 
            Height          =   300
            ItemData        =   "frmHPBR.frx":0000
            Left            =   2160
            List            =   "frmHPBR.frx":000A
            TabIndex        =   24
            Text            =   "货品名称"
            Top             =   150
            Visible         =   0   'False
            Width           =   1185
         End
      End
      Begin VB.CommandButton cmdJQ 
         Caption         =   "近期录入"
         Height          =   285
         Left            =   7920
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Frame frmYwy 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   960
         TabIndex        =   18
         Top             =   -120
         Width           =   2535
         Begin VB.TextBox txtSl 
            Height          =   270
            Left            =   630
            TabIndex        =   20
            Top             =   180
            Width           =   615
         End
         Begin VB.CommandButton cmdDao 
            BackColor       =   &H00C0C0FF&
            Caption         =   "导入"
            Height          =   315
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   180
            Width           =   765
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "数量"
            Height          =   225
            Left            =   180
            TabIndex        =   21
            Top             =   210
            Width           =   495
         End
      End
      Begin VB.CheckBox chkDel 
         Caption         =   "删除"
         Height          =   315
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   675
      End
      Begin VB.OptionButton opt2 
         Caption         =   "扩展"
         Height          =   255
         Left            =   7080
         TabIndex        =   16
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton opt1 
         Caption         =   "原始"
         Height          =   255
         Left            =   6240
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdL 
         Caption         =   "返回"
         Height          =   285
         Left            =   4110
         TabIndex        =   11
         Top             =   390
         Width           =   855
      End
      Begin VB.CommandButton cmdT 
         Caption         =   "替代产品"
         Height          =   285
         Left            =   5040
         TabIndex        =   10
         Top             =   390
         Width           =   915
      End
      Begin VB.CommandButton cmdRC 
         Caption         =   "清空表"
         Height          =   285
         Left            =   3210
         TabIndex        =   9
         Top             =   390
         Width           =   855
      End
      Begin VB.ComboBox comLx 
         Height          =   300
         ItemData        =   "frmHPBR.frx":0022
         Left            =   0
         List            =   "frmHPBR.frx":0044
         TabIndex        =   8
         Text            =   "超级搜索"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdC 
         Caption         =   "查 询"
         Height          =   285
         Left            =   2340
         TabIndex        =   7
         Top             =   390
         Width           =   825
      End
      Begin VB.TextBox txtZ 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   1185
      End
   End
   Begin VB.Frame frmBr 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5625
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgFN 
         Height          =   1695
         Left            =   7920
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2990
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame frmGx 
         Caption         =   "Frame1"
         Height          =   375
         Left            =   6120
         TabIndex        =   4
         Top             =   5280
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLP 
         Height          =   4635
         Left            =   3480
         TabIndex        =   3
         Top             =   120
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   8176
         _Version        =   393216
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdBTD 
         Caption         =   "替代"
         Height          =   255
         Left            =   6360
         TabIndex        =   2
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Timer timWait 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin VB.Timer timQuit 
         Interval        =   1000
         Left            =   630
         Top             =   90
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLN 
         Height          =   675
         Left            =   420
         TabIndex        =   1
         Top             =   750
         Visible         =   0   'False
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   1191
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgF 
         Height          =   4755
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   8387
         _Version        =   393216
         WordWrap        =   -1  'True
         SelectionMode   =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "frmHPBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bh As String
Dim B1id As Long
Dim B2id As Long
Dim B3id As Long
Dim Bm1 As String
Dim Bm2 As String
Dim Bm3 As String
Dim Bm As String
Dim HY1 As String
Dim HY2 As String
Dim HY3 As String
Dim Pid As Long
Dim timZm As Integer
Public GyId As Integer '确定双击哪个供应商
Dim Jpid As Long '近期翻页ID
Dim frId As Integer '全权限查看时的内容ID

Dim LL As String '录入者
Dim LLUid As String
Dim LCRen As String
Dim LCUid As String
Dim Lc As Integer
Dim Fwid As Long

Dim KDF As Boolean '是否可供询价单导入（被禁用，被删除的货品不能导入)
Dim LLB(4, 5000, 80) As String '多重搜索记录(列，行，次数)
Dim LCC As Integer '多重搜索到哪一步总数(对应次数）
Dim LDC As Integer '当前翻到的多重次数
Public Sub TDB(Bh As String)
Dim tt As String
Dim Ra, Rb
Dim La, Lb
Dim oo As Long
Dim R1, R2, R3, R4, R5, R6, R7
Dim JT As String
Dim LNR As String
JT = ",oname,gg,xn,pb,jz,ypb,bm1,bm2,bm3,l1,l2,l3,jyf,bz"
tt = "select bh,partname" & JT & ",pid from nlpmxc where bh='" & Bh & "' and delf=1;" & _
    "select bh,partname" & JT & ",pid from nlpmxcTdb where ybh='" & Bh & "' and delf=1 order by tid desc"
Call dtgLPFF
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Lb = UBound(Rb, 2) + 1
dtgLP.Rows = Lb + 30
dtgLN.Rows = Lb + 30
dtgLP.Visible = False
dtgLP.Row = 1
    dtgLP.Col = 0: dtgLP.Text = Ra(0, 0)
    dtgLP.Col = 1: dtgLP.Text = Ra(1, 0)
    dtgLP.Col = 2: dtgLP.Text = Ra(2, 0)
    dtgLP.Col = 3: dtgLP.Text = Ra(16, 0)
    If Ra(2, 0) <> "" Then
        LNR = "原厂编号: " & Ra(2, 0) & " " & Chr(13) & Chr(10)
    End If
    If Ra(3, 0) <> "" Then
        LNR = LNR + "包装规格: " & Ra(3, 0) & " " & Chr(13) & Chr(10)
    End If
    If Ra(4, 0) <> "" Then
        LNR = LNR + "产品型号: " & Ra(4, 0) & " " & Chr(13) & Chr(10)
    End If
    If Ra(5, 0) <> "" Then
        LNR = LNR + "适用品牌: " & Ra(5, 0) & " " & Chr(13) & Chr(10)
    End If
    If Ra(6, 0) <> "" Then
        LNR = LNR + "适用机组: " & Trim(Ra(6, 0)) & " " & Chr(13) & Chr(10)
    End If
    If Ra(7, 0) <> "" Then
        LNR = LNR + "原厂品牌: " & Ra(7, 0) & " " & Chr(13) & Chr(10)
    End If
'''''    If Ra(9, oo - 1) <> "" Then
'''''        LNR = LNR + "别名1:" & Ra(9, oo - 1) & " "
'''''    End If
'''''    If Ra(10, oo - 1) <> "" Then
'''''        LNR = LNR + "别名2:" & Ra(10, oo - 1) & " "
'''''    End If
'''''    If Ra(11, oo - 1) <> "" Then
'''''        LNR = LNR + "别名3:" & Ra(11, oo - 1) & " "
'''''    End If
'''''    If Ra(12, oo - 1) <> "" Then
'''''        LNR = LNR + "类别1:" & Ra(12, oo - 1) & " "
'''''    End If
'''''    If Ra(13, oo - 1) <> "" Then
'''''        LNR = LNR + "类别2:" & Ra(13, oo - 1) & " "
'''''    End If
'''''    If Ra(14, oo - 1) <> "" Then
'''''        LNR = LNR + "类别3:" & Ra(14, oo - 1) & " "
'''''    End If
    If Ra(15, 0) <> "" Then           '备注
        LNR = LNR + "备注: " & Ra(15, 0) & " " & Chr(13) & Chr(10)
    End If
    dtgLP.Col = 2: dtgLP.Text = LNR
    frmZu.lblDtg.Caption = LNR
    dtgLP.RowHeight(1) = frmZu.lblDtg.Height
    dtgLN.Row = 1
    dtgLN.Col = 0: dtgLN.Text = Ra(0, 0)
    dtgLN.Col = 1: dtgLN.Text = Ra(1, 0)
    dtgLN.Col = 2: dtgLN.Text = LNR
    dtgLN.Col = 3: dtgLN.Text = Ra(16, 0)
For oo = 2 To Lb + 1
    LNR = ""
    dtgLP.Row = oo
    dtgLP.Col = 0: dtgLP.Text = Rb(0, oo - 2)
    dtgLP.Col = 1: dtgLP.Text = Rb(1, oo - 2)

    dtgLP.Col = 3: dtgLP.Text = Rb(16, oo - 2)
    If Rb(2, oo - 2) <> "" Then
        LNR = "原厂编号: " & Rb(2, oo - 2) & " " & Chr(13) & Chr(10)
    End If
    If Rb(3, oo - 2) <> "" Then
        LNR = LNR + "包装规格: " & Rb(3, oo - 2) & " " & Chr(13) & Chr(10)
    End If
    If Rb(4, oo - 2) <> "" Then
        LNR = LNR + "产品型号: " & Rb(4, oo - 2) & " " & Chr(13) & Chr(10)
    End If
    If Rb(5, oo - 2) <> "" Then
        LNR = LNR + "适用品牌: " & Rb(5, oo - 2) & " " & Chr(13) & Chr(10)
    End If
    If Rb(6, oo - 2) <> "" Then
        LNR = LNR + "适用机组: " & Rb(6, oo - 2) & " " & Chr(13) & Chr(10)
    End If
    If Rb(7, oo - 2) <> "" Then
        LNR = LNR + "原厂品牌: " & Rb(7, oo - 2) & " " & Chr(13) & Chr(10)
    End If
'''''    If rb(9, oo - 1) <> "" Then
'''''        LNR = LNR + "别名1:" & rb(9, oo - 1) & " "
'''''    End If
'''''    If rb(10, oo - 1) <> "" Then
'''''        LNR = LNR + "别名2:" & rb(10, oo - 1) & " "
'''''    End If
'''''    If rb(11, oo - 1) <> "" Then
'''''        LNR = LNR + "别名3:" & rb(11, oo - 1) & " "
'''''    End If
'''''    If rb(12, oo - 1) <> "" Then
'''''        LNR = LNR + "类别1:" & rb(12, oo - 1) & " "
'''''    End If
'''''    If rb(13, oo - 1) <> "" Then
'''''        LNR = LNR + "类别2:" & rb(13, oo - 1) & " "
'''''    End If
'''''    If rb(14, oo - 1) <> "" Then
'''''        LNR = LNR + "类别3:" & rb(14, oo - 1) & " "
'''''    End If
    If Rb(15, oo - 2) <> "" Then           '备注
        LNR = LNR + "备注: " & Rb(15, oo - 2) & " " & Chr(13) & Chr(10)
    End If
    dtgLP.Col = 2: dtgLP.Text = LNR
    frmZu.lblDtg.Caption = LNR
    dtgLP.RowHeight(oo) = frmZu.lblDtg.Height
    
    dtgLN.Row = oo
    dtgLN.Col = 0: dtgLN.Text = Rb(0, oo - 2)
    dtgLN.Col = 1: dtgLN.Text = Rb(1, oo - 2)
    dtgLN.Col = 2: dtgLN.Text = LNR
    dtgLN.Col = 3: dtgLN.Text = Rb(16, oo - 2)
'''''    If oo = Lb Then
'''''        Jpid = Rb(3, oo - 1)
'''''    End If
'''''    If Jpid < 10 Then
'''''        Jpid = 0
'''''    End If
Next
dtgLP.Visible = True
dtgLP.ColWidth(2) = 6000
dtgLP.ColWidth(3) = 0
End Sub
Public Sub TD(Bh As String)
Dim tt As String
Dim Ra, Rb
Dim La, Lb
Dim oo As Long
Dim R1, R2, R3, R4, R5, R6, R7
Dim JT As String
JT = ",oname,gg,xn,pb,jz,ypb,bm1,bm2,bm3,l1,l2,l3,jyf,bz"
tt = "select bh,partname" & JT & ",pid from nlpmxc where bh='" & Bh & "' and delf=1 ;" & _
    "select bh,partname" & JT & ",pid from nlpmxcTd where ybh='" & Bh & "' and delf=1  order by tid desc"
Call dtgLPFF
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
Set mod1.HTP = mod1.HTP.NextRecordset
Rb = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
Lb = UBound(Rb, 2) + 1
dtgLP.Rows = Lb + 30
dtgLN.Rows = Lb + 30
dtgLP.Visible = False
dtgLP.Row = 1
    dtgLP.Col = 0: dtgLP.Text = Ra(0, 0)
    dtgLP.Col = 1: dtgLP.Text = Ra(1, 0)
    dtgLP.Col = 2: dtgLP.Text = Ra(2, 0)
    dtgLP.Col = 3: dtgLP.Text = Ra(3, 0)
    dtgLN.Row = 1
    dtgLN.Col = 0: dtgLN.Text = Ra(0, 0)
    dtgLN.Col = 1: dtgLN.Text = Ra(1, 0)
    dtgLN.Col = 2: dtgLN.Text = Ra(2, 0)
    dtgLN.Col = 3: dtgLN.Text = Ra(3, 0)
For oo = 2 To Lb + 1
    dtgLP.Row = oo
    dtgLP.Col = 0: dtgLP.Text = Rb(0, oo - 2)
    dtgLP.Col = 1: dtgLP.Text = Rb(1, oo - 2)
    dtgLP.Col = 2: dtgLP.Text = Rb(2, oo - 2)
    dtgLP.Col = 3: dtgLP.Text = Rb(3, oo - 2)
    dtgLN.Row = oo
    dtgLN.Col = 0: dtgLN.Text = Rb(0, oo - 2)
    dtgLN.Col = 1: dtgLN.Text = Rb(1, oo - 2)
    dtgLN.Col = 2: dtgLN.Text = Rb(2, oo - 2)
    dtgLN.Col = 3: dtgLN.Text = Rb(3, oo - 2)
'''''    If oo = Lb Then
'''''        Jpid = Rb(3, oo - 1)
'''''    End If
'''''    If Jpid < 10 Then
'''''        Jpid = 0
'''''    End If
Next
dtgLP.Visible = True
dtgLP.ColWidth(2) = 6000
dtgLP.ColWidth(3) = 0
End Sub
Public Sub dtgFF()
Dim oo As Long
dtgF.Clear: dtgFN.Clear
dtgF.Rows = 300
dtgF.Cols = 17: dtgFN.Cols = 17
dtgF.Row = 0
dtgF.Col = 0: dtgF.Text = "编号":
dtgF.CellFontBold = True
dtgF.Col = 1: dtgF.Text = "货品名称"
dtgF.CellFontBold = True
dtgF.Col = 2: dtgF.Text = "原厂编号"
dtgF.CellFontBold = True
dtgF.Col = 3: dtgF.Text = "规格型号"
dtgF.CellFontBold = True
dtgF.Col = 4: dtgF.Text = "产品型号"
dtgF.CellFontBold = True
dtgF.Col = 5: dtgF.Text = "适用品牌"
dtgF.CellFontBold = True
dtgF.Col = 6: dtgF.Text = "适用机组"
dtgF.CellFontBold = True
dtgF.Col = 7: dtgF.Text = "原厂品牌"
dtgF.CellFontBold = True
dtgF.Col = 8: dtgF.Text = "别名1"
dtgF.CellFontBold = True
dtgF.Col = 9: dtgF.Text = "别名2"
dtgF.CellFontBold = True
dtgF.Col = 10: dtgF.Text = "别名3"
dtgF.CellFontBold = True
dtgF.Col = 11: dtgF.Text = "类别1"
dtgF.CellFontBold = True
dtgF.Col = 12: dtgF.Text = "类别2"
dtgF.CellFontBold = True
dtgF.Col = 13: dtgF.Text = "类别3"
dtgF.CellFontBold = True
dtgF.Col = 14: dtgF.Text = "备注"
dtgF.CellFontBold = True
dtgF.Col = 15: dtgF.Text = "禁用否"
dtgF.CellFontBold = True
dtgF.Col = 16: dtgF.Text = "pid":
dtgF.CellFontBold = True

dtgFN.Clear
dtgFN.Rows = 300
dtgF.ColWidth(16) = 0
For oo = 1 To 15
    dtgF.ColWidth(oo) = 1500
Next
'''''For oo = 1 To 299
'''''    dtgf.RowHeight(oo) = dtgf.RowHeight(0) * 2
'''''Next
dtgF.FixedCols = 1

End Sub
Public Sub dtgLPFF()
Dim oo As Long
dtgLP.Clear
dtgLP.Rows = 300
dtgLP.Cols = 5
dtgLP.Row = 0
dtgLP.Col = 0: dtgLP.Text = "编号":
dtgLP.CellFontBold = True
dtgLP.Col = 1: dtgLP.Text = "货品名称":
dtgLP.CellFontBold = True
dtgLP.Col = 2: dtgLP.Text = "描述":
dtgLP.CellFontBold = True
dtgLP.Col = 3: dtgLP.Text = "pid":
dtgLP.CellFontBold = True
dtgLP.Col = 4: dtgLP.Text = "禁用否"
dtgLN.Clear
dtgLN.Rows = 300
dtgLN.Cols = 5
dtgLP.ColWidth(3) = 0
dtgLP.ColWidth(1) = 1860
dtgLP.ColWidth(2) = 6000
dtgLP.ColWidth(5) = 0
dtgLP.ColWidth(4) = 0
'''''For oo = 1 To 299
'''''    dtgLP.RowHeight(oo) = dtgLP.RowHeight(0) * 2
'''''Next
dtgLP.FixedCols = 1

End Sub
Public Sub dtgLPBound(tt As String)
Dim Ra
Dim La As Long
Dim Rb
Dim LNR As String
Dim zz As Integer
Dim oo As Long
Dim ii As Integer
dtgLP.Visible = False
dtgF.Visible = False
Call dtgLPFF
Call dtgFF
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2) + 1
dtgLP.Rows = La + 60
dtgLN.Rows = La + 60
dtgF.Rows = La + 60
dtgFN.Row = La + 60
    If comLx.Text = "编号" And mod1.Bm <> "技术中心" Then
        If Ra(17, 0) = False Then
        MsgBox "无此货品信息，请按其它参数进行查询或咨询技术部!"
        If opt1.Value = True Then
            dtgLP.Visible = True
        Else
            dtgF.Visible = True
        End If
        Exit Sub
        End If
    End If
For oo = 1 To La
    LNR = ""
    dtgLP.Row = oo
    dtgLP.Col = 0: dtgLP.Text = Ra(0, oo - 1)
    dtgLP.Col = 1: dtgLP.Text = Ra(1, oo - 1)
    dtgLP.Col = 3: dtgLP.Text = Ra(16, oo - 1)
    If Ra(2, oo - 1) <> "" Then
        LNR = "原厂编号: " & Ra(2, oo - 1) & " " & Chr(13) & Chr(10)
    End If
    If Ra(3, oo - 1) <> "" Then
        LNR = LNR + "包装规格: " & Ra(3, oo - 1) & " " & Chr(13) & Chr(10)
    End If
    If Ra(4, oo - 1) <> "" Then
        LNR = LNR + "产品型号: " & Ra(4, oo - 1) & " " & Chr(13) & Chr(10)
    End If
    If Ra(5, oo - 1) <> "" Then
        LNR = LNR + "适用品牌: " & Ra(5, oo - 1) & " " & Chr(13) & Chr(10)
    End If
    If Ra(6, oo - 1) <> "" Then
        LNR = LNR + "适用机组: " & Trim(Ra(6, oo - 1)) & " " & Chr(13) & Chr(10)
    End If
    If Ra(7, oo - 1) <> "" Then
        LNR = LNR + "原厂品牌: " & Ra(7, oo - 1) & " " & Chr(13) & Chr(10)
    End If
'''''    If Ra(9, oo - 1) <> "" Then
'''''        LNR = LNR + "别名1:" & Ra(9, oo - 1) & " "
'''''    End If
'''''    If Ra(10, oo - 1) <> "" Then
'''''        LNR = LNR + "别名2:" & Ra(10, oo - 1) & " "
'''''    End If
'''''    If Ra(11, oo - 1) <> "" Then
'''''        LNR = LNR + "别名3:" & Ra(11, oo - 1) & " "
'''''    End If
'''''    If Ra(12, oo - 1) <> "" Then
'''''        LNR = LNR + "类别1:" & Ra(12, oo - 1) & " "
'''''    End If
'''''    If Ra(13, oo - 1) <> "" Then
'''''        LNR = LNR + "类别2:" & Ra(13, oo - 1) & " "
'''''    End If
'''''    If Ra(14, oo - 1) <> "" Then
'''''        LNR = LNR + "类别3:" & Ra(14, oo - 1) & " "
'''''    End If
    If Ra(15, oo - 1) <> "" Then           '备注
        LNR = LNR + "备注: " & Ra(15, oo - 1) & " " & Chr(13) & Chr(10)
    End If
    dtgLP.Col = 2: dtgLP.Text = LNR
    frmZu.lblDtg.Caption = LNR
    dtgLP.RowHeight(oo) = frmZu.lblDtg.Height

    dtgLN.Row = oo
    dtgLN.Col = 0: dtgLN.Text = Ra(0, oo - 1)
    dtgLN.Col = 1: dtgLN.Text = Ra(1, oo - 1)
    dtgLN.Col = 2: dtgLN.Text = LNR
    dtgLN.Col = 3: dtgLN.Text = Ra(16, oo - 1)
    If oo = La Then
        Jpid = Ra(2, oo - 1)
    End If
    If Jpid < 10 Then
        Jpid = 0
    End If
    
    dtgF.Row = oo: dtgFN.Row = oo
    For ii = 0 To 16
        dtgF.Col = ii: dtgFN.Col = ii
        If ii = 14 Then
            dtgF.Text = Ra(ii + 1, oo - 1)
            dtgFN.Text = Ra(ii + 1, oo - 1)
        ElseIf ii = 15 Then
            If Ra(ii - 1, oo - 1) = True Then
                dtgF.Text = ""
                dtgFN.Text = ""
            Else
                dtgF.Text = "禁用"
                dtgFN.Text = "禁用"
            End If
  
        Else
            dtgF.Text = Ra(ii, oo - 1)
            dtgFN.Text = Ra(ii, oo - 1)
        End If
    Next
    
    
    
    
    
    
    
    
    
    '禁用显示红色
    If Ra(14, oo - 1) = True Then
        For zz = 0 To 16
            dtgLP.Col = zz: dtgLP.CellForeColor = &H80000012
            dtgF.Col = zz: dtgF.CellForeColor = &H80000012
        Next
    Else
        For zz = 0 To 16
            dtgLP.Col = zz: dtgLP.CellForeColor = &HFF&
            dtgF.Col = zz: dtgF.CellForeColor = &HFF&
        Next
    End If
    
    
    
    
    
Next

'查看GxBh
    If Left(Ra(0, 0), 1) = "3" Then
        tt = "select bh,partname,oname,gg,xn,pb,jz,ypb,bm1,bm2,bm3,l1,l2,l3,jyf,bz,pid,gxbh from nlpmxc where gxbh='" & Ra(0, 0) & "'"
    Else
        tt = "select bh,partname,oname,gg,xn,pb,jz,ypb,bm1,bm2,bm3,l1,l2,l3,jyf,bz,pid,gxbh from nlpmxc where bh='" & Ra(17, 0) & "'"
    End If
    Set mod1.HTP = CreateObject("adodb.recordset")
   mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    On Error Resume Next
    Rb = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    If IsNull(Rb(0, 0)) = False Then
    
                For oo = La + 1 To La + 1
                LNR = ""
                dtgLP.Row = oo
                dtgLP.Col = 0: dtgLP.Text = Rb(0, 0)
                dtgLP.Col = 1: dtgLP.Text = Rb(1, 0)
                dtgLP.Col = 3: dtgLP.Text = Rb(16, 0)
                If Rb(2, oo - 1) <> "" Then
                    LNR = "原厂编号: " & Rb(2, 0) & " " & Chr(13) & Chr(10)
                End If
                If Rb(3, oo - 1) <> "" Then
                    LNR = LNR + "包装规格: " & Rb(3, 0) & " " & Chr(13) & Chr(10)
                End If
                If Rb(4, oo - 1) <> "" Then
                    LNR = LNR + "产品型号: " & Rb(4, 0) & " " & Chr(13) & Chr(10)
                End If
                If Rb(5, oo - 1) <> "" Then
                    LNR = LNR + "适用品牌: " & Rb(5, 0) & " " & Chr(13) & Chr(10)
                End If
                If Rb(6, oo - 1) <> "" Then
                    LNR = LNR + "适用机组: " & Trim(Rb(6, 0)) & " " & Chr(13) & Chr(10)
                End If
                If Rb(7, oo - 1) <> "" Then
                    LNR = LNR + "原厂品牌: " & Rb(7, 0) & " " & Chr(13) & Chr(10)
                End If
            
                If Rb(15, oo - 1) <> "" Then           '备注
                    LNR = LNR + "备注: " & Rb(15, 0) & " " & Chr(13) & Chr(10)
                End If
                dtgLP.Col = 2: dtgLP.Text = LNR
                frmZu.lblDtg.Caption = LNR
                dtgLP.RowHeight(oo) = frmZu.lblDtg.Height
            
                dtgLN.Row = oo
                dtgLN.Col = 0: dtgLN.Text = Rb(0, 0)
                dtgLN.Col = 1: dtgLN.Text = Rb(1, 0)
                dtgLN.Col = 2: dtgLN.Text = LNR
                dtgLN.Col = 3: dtgLN.Text = Rb(16, 0)
                If oo = La Then
                    Jpid = Rb(16, 0)
                End If
                If Jpid < 10 Then
                    Jpid = 0
                End If
                
                dtgF.Row = La + 1: dtgFN.Row = La + 1
                For ii = 0 To 16
                    dtgF.Col = ii: dtgFN.Col = ii
                    If ii = 14 Then
                        dtgF.Text = Rb(ii + 1, 0)
                        dtgFN.Text = Rb(ii + 1, 0)
                    ElseIf ii = 15 Then
                        If Rb(ii - 1, 0) = True Then
                            dtgF.Text = ""
                            dtgFN.Text = ""
                        Else
                            dtgF.Text = "禁用"
                            dtgFN.Text = "禁用"
                        End If
              
                    Else
                        dtgF.Text = Rb(ii, 0)
                        dtgFN.Text = Rb(ii, 0)
                    End If
                Next
                
                
                
                
                
                
                
                
                
                '禁用显示红色
                If Ra(14, oo - 1) = True Then
                    For zz = 0 To 16
                        dtgLP.Col = zz: dtgLP.CellForeColor = &H80000012
                        dtgF.Col = zz: dtgF.CellForeColor = &H80000012
                    Next
                Else
                    For zz = 0 To 16
                        dtgLP.Col = zz: dtgLP.CellForeColor = &HFF&
                        dtgF.Col = zz: dtgF.CellForeColor = &HFF&
                    Next
                End If
                
                
                
                
                
            Next

End If
        If opt1.Value = True Then
            dtgLP.Visible = True
        Else
            dtgF.Visible = True
        End If
End Sub

Private Sub chkDel_Click()
If chkDel.Value = 1 Then
    cmdJQ.Visible = False
    'cmdT.Visible = False
ElseIf chkDel.Value = 0 Then
    cmdJQ.Visible = True
    'cmdT.Visible = True
End If
End Sub

Private Sub cmdBTD_Click()
txtZ.Text = ""
If Bh = "" Then Bh = txtZ.Text
    Call Me.TDB(Bh)
    Call frmHPZL.Qing
'清空多重搜索
Call Me.LLBQing
End Sub

Private Sub cmdC_Click()
Dim tt As String
Dim LT1 As String
Dim LT2 As String
Dim LT3 As String
Dim JT As String
Dim DelF As Integer
Dim oo As Long
Dim ii As Long
Dim zz As Integer
Dim L0 As String: Dim L1 As String: Dim L3 As String
Dim TJ As Boolean

dtgLP.Row = 1: dtgLP.Col = 0
txtZ.SelStart = 0: txtZ.SelLength = Len(txtZ.Text)
Call frmHPZL.Qing
If dtgLP.Text <> "" Then '二次查询

    If txtZ.Text = "" Then Exit Sub
    dtgLP.Clear
    dtgLP.Visible = False
    dtgLP.Row = 0
    dtgLP.Col = 0: dtgLP.Text = "编号": dtgLP.CellFontBold = True
    dtgLP.Col = 1: dtgLP.Text = "货品名称": dtgLP.CellFontBold = True
    dtgLP.Col = 2: dtgLP.Text = "描述": dtgLP.CellFontBold = True
    dtgLP.Col = 3: dtgLP.Text = "pid": dtgLP.CellFontBold = True
    dtgLP.Col = 4: dtgLP.Text = "禁用否"
'''''    dtgLP.ColWidth(2) = 0
'''''    dtgLP.ColWidth(3) = 6000
    ii = 0
    For oo = 1 To 5000
        TJ = False
        Me.dtgLN.Row = oo
        dtgLN.Col = 0
        If dtgLN.Text = "" Then Exit For
        L0 = dtgLN.Text: dtgLN.Col = 1: L1 = dtgLN.Text: dtgLN.Col = 2: L3 = dtgLN.Text
        If InStr(1, L0, txtZ.Text) > 0 Or InStr(1, L1, txtZ.Text) > 0 Or InStr(1, L3, txtZ.Text) > 0 Then
            TJ = True
            ii = ii + 1
        End If
        If TJ = True Then
            dtgLP.Row = ii
            For zz = 0 To 4
                dtgLP.Col = zz: dtgLN.Col = zz
                dtgLP.Text = dtgLN.Text
                If zz = 2 Then
                frmZu.lblDtg.Caption = dtgLN.Text
                dtgLP.RowHeight(oo) = frmZu.lblDtg.Height
                End If
            Next
        End If
        
    Next
    ' 同步内表
    dtgLN.Clear
    For oo = 1 To ii
        dtgLP.Row = oo: dtgLN.Row = oo
        For zz = 0 To 4
                dtgLP.Col = zz: dtgLN.Col = zz
                dtgLN.Text = dtgLP.Text
        Next
    Next
    dtgLP.Visible = True
    '追加记录
    LCC = LCC + 1
    LDC = LDC + 1
    For oo = 0 To 5000
        dtgLN.Row = oo + 1
        dtgLN.Col = 0
        If dtgLN.Text = "" Then Exit For
        For zz = 0 To 4
            dtgLN.Col = zz
            LLB(zz, oo, LDC) = dtgLN.Text
        Next
    Next
    Exit Sub
End If
DelF = 1
If chkDel.Value = 1 Then
    DelF = 0
End If
JT = ",oname,gg,xn,pb,jz,ypb,bm1,bm2,bm3,l1,l2,l3,jyf,bz"
Select Case comLx.Text
Case "超级搜索"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bh='" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "' or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " bz like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%')  and delf=" & DelF & " "
Case "编号"
    If Len(Replace(txtZ.Text, vbCrLf, "", 1)) = 1 And Val(Replace(txtZ.Text, vbCrLf, "", 1)) > 0 Then
        tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where left(bh,1)='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=" & DelF & " "
    ElseIf Len(Replace(txtZ.Text, vbCrLf, "", 1)) = 2 And Val(Replace(txtZ.Text, vbCrLf, "", 1)) > 0 Then
        tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where left(bh,2)='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=" & DelF & " "
    ElseIf Len(Replace(txtZ.Text, vbCrLf, "", 1)) = 3 And Val(Replace(txtZ.Text, vbCrLf, "", 1)) > 0 Then
        tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where left(bh,3)='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=" & DelF & " "
    ElseIf Len(Replace(txtZ.Text, vbCrLf, "", 1)) = 5 And Val(Replace(txtZ.Text, vbCrLf, "", 1)) > 0 Then
        tt = "select bh,partname" & JT & ",pid,gxbh,delf from nlpmxc where bh='" & Replace(txtZ.Text, vbCrLf, "", 1) & "'"
    End If
Case "类别"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where (l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%')  and delf=" & DelF
Case "别名"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where (bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%')  and delf=" & DelF
Case "原厂编号"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where oname='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=" & DelF
Case "适用品牌"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where pb='" & Replace(txtZ.Text, vbCrLf, "", 1) & "' and delf=" & DelF & " "
Case "适用机组"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where jz like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' and delf=" & DelF & " "
Case "分类"
    tt = "select bh,partname,'原厂编号:'+oname+' '+gg+' '+xn+' ',pid,gxbh from nlpmxc where (lb1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or lb2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%')  and delf=" & DelF & " "
Case "豪曼产品"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where left(bh,1)='H' and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') and delf=" & DelF & " "
Case "特殊替代类"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where left(bh,1)='B' and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') and delf=" & DelF & "  "
Case "事后流程及易耗"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where left(bh,1)='A' and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') and delf=" & DelF & "  "
Case "原厂零件"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where left(bh,1)='9' and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') and delf=" & DelF & "  "
Case "产品类"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where left(bh,1)='8' and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') and delf=" & DelF & "  "
Case "临时类"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where left(bh,1)='3' and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') "
Case "替代类"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where (left(bh,1)='1'" & _
        "or left(bh,1)='2' or left(bh,1)='4' or left(bh,1)='5' or left(bh,1)='6' or left(bh,1)='7') and (partname like '%" & _
    Replace(txtZ.Text, vbCrLf, "", 1) & "%' or oname like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%'  or ypb like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or jz like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%'  or xn like  '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & _
    "%' or bm2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or bm3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l1 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or l2 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%' or" & _
    " l3 like '%" & Replace(txtZ.Text, vbCrLf, "", 1) & "%') and delf=" & DelF & "  "
End Select
If tt = "" Then Exit Sub
If mod1.Bm <> "技术中心" And mod1.Bm <> "维保中心" And mod1.DName <> "货品录入员" Then
    'If comLx.Text <> "编号" And Left(txtZ.Text, 1) <> "3") Then
        tt = tt & " and (lc=100 or left(bh,1)='3')"
    'End If
End If
tt = tt & " order by bh"
Call dtgLPBound(tt)
'清空多重搜索
Call Me.LLBQing
'''''dtgLP.ColWidth(2) = 0
'''''dtgLP.ColWidth(3) = 6000
'cmdT.Visible = False

End Sub

Private Sub cmdDao_Click()
Dim hg As Long
On Error Resume Next
'If comLx.Text = "" Then Exit Sub
Dim tt As String
Dim Ra
tt = "select delf,jyf from nlpmxc where bh='" & Bh & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
On Error Resume Next
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
If Text1.Text = "" Then
If (Ra(0, 0) = False Or Ra(1, 0) = False) Then
    MsgBox ("已删除或禁用的货品不支持导入!")
    Exit Sub
End If
End If
'''''''If KDF = False Then
'''''''    MsgBox ("已删除或禁用的货品不支持导入!")
'''''''    Exit Sub
'''''''End If
    If Val(txtSl.Text) = 0 Then
        MsgBox "请确认数量!"
        txtSl.SetFocus
        Exit Sub
    End If
    If FmxcXJ.txtDRQ.Text = "" Then
        FmxcXJ.txtDRQ.Text = mod1.DQda
    End If
    If FmxcXJ.txtBrq.Text = "" Then
        FmxcXJ.txtBrq.Text = mod1.DQda
    End If


    If Bh = "" And Text1.Text = "" Then Exit Sub
    
If FmxcXJ.Visible = True Then
    
                                      '新版本速达
        timZm = 2
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "询价单2011"
        mod1.cmd.Parameters("@NBLX") = "豪曼配件添加"
        mod1.cmd.Parameters("@bh") = ""
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = FmxcXJ.lblBid.ToolTipText
        mod1.cmd.Parameters("@mt2") = FmxcXJ.lblZl.Caption
        mod1.cmd.Parameters("@mt3") = FmxcXJ.txtLx.Text '业务类型
        mod1.cmd.Parameters("@mt7") = Bh

        mod1.cmd.Parameters("@mlt1") = Text1.Text '人工其它的添加
        mod1.cmd.Parameters("@mm1") = Val(txtSl.Text) '数量
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
ElseIf fmxcZJ.Visible = True Then
        timZm = 3
        Set mod1.cmd = CreateObject("adodb.command")
        mod1.cmd.ActiveConnection = mod1.cc
        mod1.cmd.CommandText = "MLAdd"
        mod1.cmd.CommandType = adCmdStoredProc
        mod1.cmd.Parameters("@zid") = 0
        mod1.cmd.Parameters("@errch") = ""
        mod1.cmd.Parameters("@NB") = "成本追加单"
        mod1.cmd.Parameters("@NBLX") = "新货品添加"
        mod1.cmd.Parameters("@bh") = fmxcZJ.lblZid.ToolTipText
        mod1.cmd.Parameters("@ywy") = mod1.DName
        mod1.cmd.Parameters("@uid") = mod1.DHid
        mod1.cmd.Parameters("@mt1") = ""
        If fmxcZJ.Visible = True And cmdDao.Caption = "分包导入" Then
            mod1.cmd.Parameters("@mt7").Value = "分包"
        Else
            mod1.cmd.Parameters("@mt7") = Bh
        End If
        mod1.cmd.Parameters("@mlt1") = ""
        mod1.cmd.Parameters("@mm1") = Val(txtSl.Text) '数量
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
End If
End Sub

Private Sub cmdGB_Click()
Me.Visible = False
If frmHPZL.Visible = True Then
    frmHPZL.Enabled = True
    frmHPZL.Show
    frmHPZL.ZOrder 0
ElseIf FmxcXJ.Visible = True Then
    FmxcXJ.Enabled = True
    FmxcXJ.Show
    FmxcXJ.ZOrder 0
End If
frmZu.Enabled = True
End Sub

Private Sub cmdJQ_Click()
Dim tt As String
Dim JT As String
JT = ",oname,gg,xn,pb,jz,ypb,bm1,bm2,bm3,l1,l2,l3"
If Jpid = 0 Then
    tt = "select top 50 bh,partname,pid" & JT & " from nlpmxc where delf=1  order by pid desc"
Else
    tt = "select top 50 bh,partname,pid" & JT & " from nlpmxc where delf=1 and pid<=" & Jpid & " order by pid desc"
End If
Call dtgLPBound(tt)
End Sub

Private Sub cmdL_Click()
Dim oo As Long
Dim ii As Integer
CCT:
If LDC = LCC Then
    LDC = LDC - 2
Else
    LDC = LDC - 1
End If
If LDC = -1 Then
    LDC = 0
    Exit Sub
End If
Call dtgLPFF

dtgLP.Visible = False
For oo = 0 To 5000
    If LLB(0, oo, LDC) = "" And oo > 0 Then Exit For
    dtgLP.Rows = oo + 30
    dtgLN.Rows = oo + 30
    dtgLP.Row = oo + 1
    dtgLN.Row = oo + 1

    For ii = 0 To 4
        If ii = 0 Then
            If LLB(ii, oo, LDC) = "" Then
                GoTo CCT
            End If
        End If
        dtgLP.Col = ii: dtgLN.Col = ii
        dtgLP.Text = LLB(ii, oo, LDC)
        dtgLN.Text = LLB(ii, oo, LDC)
        If ii = 2 Then
            frmZu.lblDtg.Caption = dtgLN.Text
            dtgLP.RowHeight(oo + 1) = frmZu.lblDtg.Height
        End If
    Next
Next
dtgLP.Visible = True
End Sub

Private Sub cmdR_Click()
Dim oo As Long
Dim ii As Integer
LDC = LDC + 1
If LDC > LCC Then
    LDC = LCC
    Exit Sub
End If
Call dtgLPFF

dtgLP.Visible = False
For oo = 0 To 5000
    If LLB(0, oo, LDC) = "" Then Exit For
    dtgLP.Rows = oo + 30
    dtgLN.Rows = oo + 30
    dtgLP.Row = oo + 1
    dtgLN.Row = oo + 1

    For ii = 0 To 4
        dtgLP.Col = ii: dtgLN.Col = ii
        dtgLP.Text = LLB(ii, oo, LDC)
        dtgLN.Text = LLB(ii, oo, LDC)
        If ii = 2 Then
            frmZu.lblDtg.Caption = dtgLN.Text
            dtgLP.RowHeight(oo + 1) = frmZu.lblDtg.Height
        End If
    Next
Next
dtgLP.Visible = True
End Sub


Private Sub cmdRC_Click()
Call Me.dtgLPFF
Call Me.dtgFF
Call frmHPZL.Qing
''''''Dim oo As Long
''''''Dim ii As Long
''''''Dim zz As Integer
''''''Dim l0 As String: Dim L1 As String: Dim L3 As String
''''''Dim TJ As Boolean
''''''If txtZ.Text = "" Then Exit Sub
''''''dtgLP.Clear
''''''dtgLP.Visible = False
''''''dtgLP.Row = 0
''''''dtgLP.Col = 0: dtgLP.Text = "编号": dtgLP.CellFontBold = True
''''''dtgLP.Col = 1: dtgLP.Text = "货品名称": dtgLP.CellFontBold = True
''''''dtgLP.Col = 3: dtgLP.Text = "描述": dtgLP.CellFontBold = True
''''''dtgLP.Col = 2: dtgLP.Text = "描述": dtgLP.CellFontBold = True
''''''dtgLP.Col = 4: dtgLP.Text = "禁用否"
''''''dtgLP.ColWidth(2) = 0
''''''dtgLP.ColWidth(3) = 6000
''''''ii = 0
''''''For oo = 1 To 5000
''''''    TJ = False
''''''    Me.dtgLN.Row = oo
''''''    dtgLN.Col = 0
''''''    If dtgLN.Text = "" Then Exit For
''''''    l0 = dtgLN.Text: dtgLN.Col = 1: L1 = dtgLN.Text: dtgLN.Col = 3: L3 = dtgLN.Text
''''''    If InStr(1, l0, txtZ.Text) > 0 Or InStr(1, L1, txtZ.Text) > 0 Or InStr(1, L3, txtZ.Text) > 0 Then
''''''        TJ = True
''''''        ii = ii + 1
''''''    End If
''''''    If TJ = True Then
''''''        dtgLP.Row = ii
''''''        For zz = 0 To 4
''''''            dtgLP.Col = zz: dtgLN.Col = zz
''''''            dtgLP.Text = dtgLN.Text
''''''        Next
''''''    End If
''''''
''''''Next
''''''' 同步内表
''''''dtgLN.Clear
''''''For oo = 1 To ii
''''''    dtgLP.Row = oo: dtgLN.Row = oo
''''''    For zz = 0 To 4
''''''            dtgLP.Col = zz: dtgLN.Col = zz
''''''            dtgLN.Text = dtgLP.Text
''''''    Next
''''''Next
''''''dtgLP.Visible = True
End Sub

Private Sub cmdT_Click()
txtZ.Text = ""
If Bh = "" Then Bh = txtZ.Text
    Call Me.TDB(Bh)
    Call frmHPZL.Qing
'清空多重搜索
Call Me.LLBQing

End Sub


Private Sub cmdXT_Click()
Dim tt As String
Dim JT As String
JT = ",oname,gg,xn,pb,jz,ypb,bm1,bm2,bm3,l1,l2,l3"
Select Case comMLx.Text
Case "货品名称"
    tt = "selectt bh,partname,pid" & JT & " from nlpmxc" & _
    " where partname in (select partname from nlpmxc group by partname having(count(*))>1)"
    
Case "适用机组"
    tt = "selectt bh,partname,pid" & JT & " from nlpmxc" & _
    " where jz in (select jz from nlpmxc group by jz having(count(*))>1)"
Case ""
    tt = "selectt bh,partname,pid" & JT & " from nlpmxc order by pb desc,partname,jz"
End Select
Call dtgLPBound(tt)
End Sub

Private Sub Command1_Click()
Dim JT As String
Dim tt As String
JT = ",oname,gg,xn,pb,jz,ypb,bm1,bm2,bm3,l1,l2,l3,jyf,bz"
    tt = "select bh,partname" & JT & ",pid,gxbh from nlpmxc where ll='导入' and lc=1 order by pid"


Call dtgLPBound(tt)
End Sub

Private Sub Command2_Click()
'Call NewBh(dtgN1.Text)
End Sub


Private Sub dtgF_Click()
On Error Resume Next



dtgFN.Row = dtgF.Row
dtgFN.Col = 16
Pid = Val(dtgFN.Text)

    dtgFN.Col = 0
    Bh = dtgFN.Text
If Pid = 0 Then Exit Sub
If mod1.Bm <> "技术中心" And mod1.Bm <> "市场营销部" And mod1.Bm <> "维保中心" And mod1.DName <> "货品录入员" And mod1.DName <> "邹晨" Then Exit Sub
'''''If txtTdbh.Visible = False Then
    Call frmHPZL.Qing
    Call frmHPZL.Bound(Pid)
'''''Else
'''''    dtgLN.Col = 0
'''''    Bh = dtgLN.Text
'''''    If Bh <> txtBh.Text Then
'''''        txtTdbh.Text = Bh
'''''    End If
'''''End If
cmdT.Visible = True
'''''If Left(Bh, 1) = "9" Then
'''''    cmdT.Caption = "替代产品"
'''''Else
'''''    cmdT.Caption = "替代"
'''''End If

If dtgF.ForeColor = &H80000012 Then
    KDF = False
Else
    KDF = True
End If
End Sub

Private Sub dtgLP_Click()

On Error Resume Next


frmHPBR.Enabled = True
dtgLN.Row = dtgLP.Row
dtgLN.Col = 3
Pid = Val(dtgLN.Text)

    dtgLN.Col = 0
    Bh = dtgLN.Text
If Pid = 0 Then Exit Sub
If mod1.Bm <> "技术中心" And mod1.Bm <> "市场营销部" And mod1.DName <> "倪东海" And mod1.DName <> "货品录入员" And mod1.DName <> "邹晨" And mod1.DName <> "李午阳" Then Exit Sub
'''''If txtTdbh.Visible = False Then
    Call frmHPZL.Qing
    Call frmHPZL.Bound(Pid)
'''''Else
'''''    dtgLN.Col = 0
'''''    Bh = dtgLN.Text
'''''    If Bh <> txtBh.Text Then
'''''        txtTdbh.Text = Bh
'''''    End If
'''''End If
cmdT.Visible = True
'''''If Left(Bh, 1) = "9" Then
'''''    cmdT.Caption = "替代产品"
'''''Else
'''''    cmdT.Caption = "替代"
'''''End If

If dtgLP.ForeColor = &H80000012 Then
    KDF = False
Else
    KDF = True
End If
End Sub

Private Sub Form_Load()
Dim tt As String
Dim Ra: Dim Rb: Dim RC: Dim RD
Me.dtgLP.Left = 0
dtgLP.Top = 0
Me.Width = 10980
Me.Height = 6675

If FmxcXJ.Visible = True Or fmxcZJ.Visible = True Then
    frmYwy.Visible = True
Else
    frmYwy.Visible = False
End If

tt = "select count(pid) from nlpmxc;select count(pid) from nlpmxc where delf=1 and jyf=1 and lc=100;" & _
    "select count(pid) from nlpmxc where jyf=0;select count(pid) from nlpmxc where delf=0"
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
mod1.HTP.Close
Set mod1.HTP = Nothing
Me.Caption = "货品资料查询　　已录入货品总数：" & Ra(0, 0) & "  有效货品数：" & Rb(0, 0) & "  禁用货品数：" & RC(0, 0) & "  删除货品数：" & RD(0, 0)
If mod1.DName <> "倪东海" And mod1.DName <> "马晓聪" Then
    cmdBTD.Visible = False
    chkDel.Visible = False
Else
    cmdJQ.Visible = True
    frmXT.Visible = True
    chkDel.Visible = True
End If
If mod1.DName = "倪东海" Or mod1.DName = "李午阳" Or mod1.DName = "邹晨" Or mod1.DName = "马晓聪" Or mod1.DName = "货品录入员" Then
    frmDao.Visible = True
    
Else
    frmDao.Visible = False
End If
End Sub

Private Sub Form_Resize()
frmBr.Width = Me.Width
dtgLP.Width = Me.Width - 200
frmBr.Height = Me.Height - 2000
dtgLP.Height = Me.Height - 2000
frmC.Top = Me.Height - 1800
cmdGB.Left = Me.Width - 1000
cmdGB.Top = Me.Height - 1000

dtgF.Height = Me.Height - 2000
dtgF.Top = 0
dtgF.Width = Me.Width - 200
dtgF.Left = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
If frmHPZL.Visible = True Then
    frmHPZL.Enabled = True
    frmHPZL.Show
    frmHPZL.ZOrder 0
ElseIf FmxcXJ.Visible = True Then
    FmxcXJ.Enabled = True
    FmxcXJ.Show
    FmxcXJ.ZOrder 0
End If
End Sub

Private Sub opt1_Click()
dtgLP.Visible = True
dtgF.Visible = False
End Sub

Private Sub opt2_Click()
dtgLP.Visible = False
dtgF.Visible = True
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
txtSl.Text = 1
txtZ.Text = ""
Bh = ""
End Sub

Private Sub timQuit_Timer()
Dim tt As String
Dim Rb, RC, RD, RE
Dim Lb As Integer
On Error Resume Next
Dim oo As Integer
Dim jj As Integer
Unload frmWaitA
Me.Enabled = True
mod1.Ti = 0
If timZm = 2 Then
    tt = "select ljbh,detail,mj,dj,jdj,sl,jhg,drq,zbq,delf,lid,ljmc,gyid1,gyid2,gyid3,gdj1,gdj2,gdj3,mc1,mc2,mc3,gyid,sddj,sdxg,sdyh,ywlx  from XJDetail where bid=" & Val(FmxcXJ.lblBid.ToolTipText) & " order by delf desc,lid desc"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    Rb = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Lb = UBound(Rb, 2)
    Call FmxcXJ.dtgBrBound(Rb, Lb)
ElseIf timZm = 3 Then
    tt = "declare @hid int;" & _
        "select @hid=hid from htzui where zid=" & Val(fmxcZJ.lblZid.ToolTipText) & ";" & _
    "select bh,nr,dj,jdj,sl,ze,delf,did,gyid1,gyid2,gyid3,gdj1,gdj2,gdj3,mc1,mc2,mc3,gyid,sddj,sdxg  from zuijiaDetail where zid=" & Val(fmxcZJ.lblZid.ToolTipText) & " order by delf desc,did desc;" & _
            "select sum(ze) from htzuidetail where zid=" & Val(fmxcZJ.lblZid.ToolTipText) & ";" & _
        "select sum(ze) from htzuiZe where hid=@hid"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
    RC = mod1.HTP.GetRows
    Set mod1.HTP = mod1.HTP.NextRecordset
    RD = mod1.HTP.GetRows
    Set mod1.HTP = mod1.HTP.NextRecordset
    RE = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    Call fmxcZJ.NewMxBound(RC, RD, RE)
End If
timQuit.Enabled = False
End Sub

Private Sub timWait_Timer()
Dim tt As String
Dim ii As Integer
On Error Resume Next
timWait.Enabled = False

tt = "select cf,bz,bh,mm1,mm2,mt1,mt2 from ml where zid=" & mod1.Zid
Set mod1.WP = CreateObject("adodb.recordset")
mod1.WP.Open tt, mod1.workBh, adOpenKeyset, adLockReadOnly, adCmdText
If mod1.WP.Fields("cf").Value = 1 Then '提交成功
    frmWaitA.lblRun.Width = frmWaitA.Width
    timQuit.Enabled = True
    mod1.Ti = 0

    timWait.Enabled = False
    Exit Sub
ElseIf mod1.WP.Fields("cf").Value = 0 And mod1.Ti < 5 Then '未完成

ElseIf mod1.WP.Fields("cf").Value = 2 Then  '处理失败
    ii = MsgBox("服务中心在处理您的命令时,发生如下错误:" & Chr(13) & mod1.WP.Fields("bz").Value, vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
''''''''    If timZm = 1 Then
''''''''        cmdJG.Enabled = False
''''''''    End If
    timWait.Enabled = False
    Exit Sub
ElseIf mod1.Ti = 5 Then
    ii = MsgBox("服务中心在处理您的命令时,超时!", vbExclamation + vbOKOnly, "二级警告!")
    Unload frmWaitA
    Me.Enabled = True
    mod1.Ti = 0
'''''    If timZm = 1 Then
'''''        cmdJG.Enabled = False
'''''    End If
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


Public Sub LLBQing()
Dim oo As Integer
Dim ii As Integer
Dim zz As Long
'''''For zz = 0 To 20
'''''    For oo = 0 To 5000
'''''        For ii = 0 To 4
'''''            LLB(ii, oo, zz) = ""
'''''        Next
'''''    Next
'''''Next

For oo = 0 To 5000
    dtgLN.Row = oo + 1
    dtgLN.Col = 0
    If dtgLN.Text = "" Then
        Exit For
    End If
    For ii = 0 To 4
        dtgLN.Col = ii
        LLB(ii, oo, LCC) = dtgLN.Text
    Next
Next
LCC = LCC + 1
LDC = LCC
End Sub
