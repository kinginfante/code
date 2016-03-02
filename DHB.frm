VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form DHB 
   Caption         =   "单位电话"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   15210
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdBao 
      BackColor       =   &H008080FF&
      Caption         =   "保存"
      Height          =   465
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7950
      Width           =   6105
   End
   Begin RichTextLib.RichTextBox text1 
      Height          =   7965
      Left            =   9060
      TabIndex        =   28
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   14049
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"DHB.frx":0000
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgLL 
      Height          =   615
      Left            =   5520
      TabIndex        =   27
      Top             =   8070
      Visible         =   0   'False
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame frmED 
      Caption         =   "编辑"
      Height          =   1125
      Left            =   120
      TabIndex        =   14
      Top             =   6870
      Width           =   8925
      Begin VB.CommandButton cmdClose 
         Caption         =   "关闭"
         Height          =   405
         Left            =   8280
         TabIndex        =   26
         Top             =   660
         Width           =   585
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "保存"
         Height          =   405
         Left            =   7530
         TabIndex        =   25
         Top             =   660
         Width           =   675
      End
      Begin VB.TextBox txtPhoX 
         Height          =   270
         Left            =   3870
         TabIndex        =   24
         Top             =   645
         Width           =   1035
      End
      Begin VB.TextBox txtUserpho 
         Height          =   270
         Left            =   5070
         TabIndex        =   23
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtPhoD 
         Height          =   270
         Left            =   1080
         TabIndex        =   22
         Top             =   675
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "内部号码"
         Height          =   285
         Left            =   2910
         TabIndex        =   21
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "手机"
         Height          =   225
         Left            =   4440
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "单位电话"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   675
         Width           =   915
      End
      Begin VB.Label lblUid 
         Height          =   255
         Left            =   3570
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblRen 
         Height          =   255
         Left            =   1110
         TabIndex        =   17
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label2 
         Caption         =   "工号"
         Height          =   255
         Left            =   2910
         TabIndex        =   16
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "姓名"
         Height          =   255
         Left            =   540
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frmZzf 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   30
      TabIndex        =   10
      Top             =   7920
      Width           =   2895
      Begin VB.OptionButton opt1 
         Caption         =   "在职"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   150
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton opt3 
         Caption         =   "离职"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   12
         Top             =   150
         Width           =   855
      End
      Begin VB.CheckBox chkHH 
         Caption         =   "转证"
         Height          =   255
         Left            =   2220
         TabIndex        =   11
         Top             =   210
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdRef 
      Caption         =   "查   询"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7920
      TabIndex        =   9
      Top             =   8670
      Width           =   1095
   End
   Begin VB.Frame frmTj 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   8580
      Width           =   6555
      Begin VB.ComboBox txtZ 
         Height          =   300
         ItemData        =   "DHB.frx":009D
         Left            =   4500
         List            =   "DHB.frx":009F
         TabIndex        =   5
         Top             =   120
         Width           =   1965
      End
      Begin VB.ComboBox comLx 
         Height          =   300
         ItemData        =   "DHB.frx":00A1
         Left            =   990
         List            =   "DHB.frx":00AE
         TabIndex        =   4
         Top             =   150
         Width           =   1485
      End
      Begin VB.ComboBox comBj 
         Height          =   300
         ItemData        =   "DHB.frx":00C4
         Left            =   3120
         List            =   "DHB.frx":00CE
         TabIndex        =   3
         Text            =   "="
         Top             =   150
         Width           =   825
      End
      Begin VB.Label Label5 
         Caption         =   "查询类别:"
         Height          =   255
         Left            =   30
         TabIndex        =   8
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "值:"
         Height          =   255
         Left            =   4110
         TabIndex        =   7
         Top             =   180
         Width           =   315
      End
      Begin VB.Label Label4 
         Caption         =   "比较:"
         Height          =   225
         Left            =   2610
         TabIndex        =   6
         Top             =   180
         Width           =   585
      End
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "全部"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6810
      TabIndex        =   1
      Top             =   8670
      Width           =   945
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgDHB 
      Height          =   7995
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   14102
      _Version        =   393216
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "DHB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ott As String
Public TC As String
Private Sub cmdED_Click()
frmED.Visible = True
End Sub

Private Sub chkAll_Click()
If chkAll.Value = 1 Then
    frmTj.Enabled = False
    opt1.Value = False
    opt3.Value = False
    frmZzf.Enabled = False
Else
    frmTj.Enabled = True
    comLx.Text = ""
    txtZ.Text = ""
    opt1.Value = True
    frmZzf.Enabled = True
End If
End Sub

Private Sub cmdBao_Click()
Dim tt As String
On Error GoTo DHBERR5
tt = "update gr set gr='" & text1.Text & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = Nothing

Exit Sub
DHBERR5:
MsgBox "出错!"
End Sub

Private Sub cmdClose_Click()
frmED.Visible = False
End Sub

Private Sub cmdOK_Click()
Dim tt As String
Dim oo As Integer: Dim ii As Integer
Dim Ra: Dim La
On Error Resume Next
tt = "update worker set userpho='" & txtUserpho.Text & "',phod='" & txtPhoD.Text & "',phox='" & txtPhoX.Text & "' where userid='" & lblUid.Caption & "';" & Ott
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workFF, adOpenForwardOnly, adLockReadOnly, adCmdText
Set mod1.HTP = mod1.HTP.NextRecordset
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
La = UBound(Ra, 2)
dtgDHB.Rows = La + 30
For oo = 1 To La + 1
    dtgDHB.Row = oo
    For ii = 1 To 10
        dtgDHB.Col = ii
        dtgDHB.Text = Ra(ii - 1, oo - 1)
    Next
Next

dtgDHB.MergeCol(1) = True
dtgDHB.MergeCells = 3
'''''''frmED.Visible = False
Call Bound(TC)
End Sub

Private Sub cmdRef_Click()
Dim oo As Integer
Dim ii As Integer
Dim Ra: Dim La
Dim tt As String
Dim Qy As String
Dim ZZF As Integer
Dim hgF As Integer
On Error Resume Next
dtgDHB.Visible = False
dtgDHB.Clear
dtgLL.Clear
dtgDHB.Row = 0
dtgDHB.Col = 1: dtgDHB.Text = "部门"
dtgDHB.Col = 2: dtgDHB.Text = "姓名": dtgDHB.Col = 3: dtgDHB.Text = "单位电话": dtgDHB.Col = 4: dtgDHB.Text = "手机": dtgDHB.Col = 5: dtgDHB.Text = "内部号码"

If opt1.Value = True Then
    ZZF = 1
Else
    ZZF = 0
End If

If chkAll.Value = 1 Then
        tt = "select bm,username,phoD,userPho,phoX,userid from renyuan1 where username<>'匿名者'  order by bmid,userid"
Else

    Select Case comLx.Text
    Case "区域"
        tt = "select bm,username,phoD,userPho,phoX,userid from renyuan1 where zzf=" & ZZF & " and qy='" & txtZ.Text & "' order by bmid,userid"
    Case "部门"

            tt = "select bm,username,phoD,userPho,phoX,userid from renyuan1 where zzf=" & ZZF & " and bm='" & txtZ.Text & "' order by userid"
            If txtZ.Text = "行政人事" Then
                tt = "select bm,username,phoD,userPho,phoX,userid from renyuan1 where zzf=" & ZZF & " and bm='" & txtZ.Text & "' or bm='商务部' order by bmid,userid"
            End If
        
    
    Case "姓名"

            tt = "select bm,username,phoD,userPho,phoX,userid from renyuan1 where zzf=" & ZZF & " and username like '%" & txtZ.Text & "%' order by userid"
    Case Else '没有条件查询，只有在离职的查询
        tt = "select bm,username,phoD,userPho,phoX,userid from renyuan1 where username<>'匿名者' and zzf=" & ZZF & " order by bmid,userid"
    End Select
End If
TC = tt
''''''''''''''Ott = TT
''''''''''''''Set mod1.HTP = CreateObject("adodb.recordset")
''''''''''''''mod1.HTP.Open TT, mod1.workFF, adOpenForwardOnly, adLockReadOnly, adCmdText
''''''''''''''Ra = mod1.HTP.GetRows
''''''''''''''mod1.HTP.Close
''''''''''''''Set mod1.HTP = Nothing
''''''''''''''La = UBound(Ra, 2)
''''''''''''''dtgDHB.Rows = La + 30
''''''''''''''Dim OBm As String
''''''''''''''For oo = 1 To La + 1
''''''''''''''    dtgDHB.Row = oo
''''''''''''''    dtgLL.Row = oo
''''''''''''''    For ii = 1 To 10
''''''''''''''        dtgDHB.Col = ii: dtgLL.Col = ii
''''''''''''''        dtgDHB.Text = Ra(ii - 1, oo - 1)
''''''''''''''        dtgLL.Text = dtgDHB.Text
''''''''''''''        If ii = 1 Then
''''''''''''''            If OBm <> dtgDHB.Text Then
''''''''''''''                OBm = dtgDHB.Text
''''''''''''''            Else
''''''''''''''                dtgDHB.Text = ""
''''''''''''''            End If
''''''''''''''        End If
''''''''''''''    Next
''''''''''''''Next
Call Bound(tt)
''''''dtgDHB.MergeCol(1) = True
''''''dtgDHB.MergeCells = 3
dtgDHB.Visible = True
End Sub

Private Sub comLx_Click()
Dim adoZ As Object
Dim oo As Integer
Dim tt As String
Dim Ra: Dim La

On Error Resume Next
Select Case comLx.Text
Case "区域"
    For oo = 20 To 0 Step -1
        txtZ.RemoveItem oo
    Next
    tt = "select qy from yzqy order by qid"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workFF, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2)
    For oo = 0 To La
        txtZ.AddItem Ra(0, oo)
    Next

    comBj.Text = "="
Case "部门"
    For oo = 20 To 0 Step -1
        txtZ.RemoveItem oo
    Next
    tt = "select bm from renyuan1 where zzf=1 group by bm,bmid,zzf order by bmid"
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workFF, adOpenForwardOnly, adLockReadOnly, adCmdText
    Ra = mod1.HTP.GetRows
    mod1.HTP.Close
    Set mod1.HTP = Nothing
    La = UBound(Ra, 2)
    For oo = 0 To La
        txtZ.AddItem Ra(0, oo)
    Next
    comBj.Text = "="
Case "姓名"
    comBj.Text = "包含"
End Select
txtZ.Text = ""
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()

End Sub

Private Sub dtgDHB_Click()
Call BQing
On Error Resume Next
dtgLL.Row = dtgDHB.Row
dtgLL.Col = dtgDHB.Col

dtgLL.Col = 2
lblRen.Caption = dtgLL.Text
dtgLL.Col = 6
lblUid.Caption = dtgLL.Text
dtgLL.Col = 3
txtPhoD.Text = dtgLL.Text
dtgLL.Col = 4
txtUserpho.Text = dtgLL.Text
dtgLL.Col = 5
txtPhoX.Text = dtgLL.Text


End Sub

Private Sub dtgDHB_DblClick()
If mod1.DName = "陈珊珊" Or mod1.DName = "吴之禺" Or mod1.DName = "李莉娜" Then
    frmED.Visible = True
End If
End Sub


Private Sub Form_Load()
Dim tt As String
Dim Ra
dtgDHB.Cols = 9
dtgDHB.Row = 0
dtgDHB.Col = 1: dtgDHB.Text = "部门"
dtgDHB.Col = 2: dtgDHB.Text = "姓名": dtgDHB.Col = 3: dtgDHB.Text = "单位电话": dtgDHB.Col = 4: dtgDHB.Text = "手机": dtgDHB.Col = 5: dtgDHB.Text = "内部号码"

dtgDHB.ColWidth(0) = 300
dtgDHB.ColWidth(1) = 1500
dtgDHB.ColWidth(3) = 1500
dtgDHB.ColWidth(4) = 2500
dtgDHB.ColWidth(5) = 1500
dtgDHB.ColWidth(6) = 0
frmED.Visible = False
frmED.Left = 0
frmED.Top = 6870
Me.Left = 0
Me.Top = 0
If mod1.DName = "陈珊珊" Or mod1.DName = "吴之禺" Or mod1.DName = "陆俊洁" Or mod1.DName = "黄雅琴" Or mod1.DName = "马晓聪" Then
    frmED.Visible = True
Else
    cmdBao.Visible = False
End If

tt = "select gr from gr"
Set mod1.HTP = Nothing
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenForwardOnly, adLockReadOnly, adCmdText
Ra = mod1.HTP.GetRows
mod1.HTP.Close
Set mod1.HTP = Nothing
text1.Text = Ra(0, 0)
End Sub

Public Sub BQing()
lblRen.Caption = ""
lblUid.Caption = ""
txtPhoD.Text = ""
txtUserpho.Text = ""
txtPhoX.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmZu.TBa.Buttons(7).Value = tbrUnpressed
End Sub


Public Sub Bound(tt As String)
Dim Ra
Dim La As Integer
Dim OBm As String
Dim oo As Integer
Dim ii As Integer
        Set mod1.HTP = CreateObject("adodb.recordset")
        mod1.HTP.Open tt, mod1.workFF, adOpenForwardOnly, adLockReadOnly, adCmdText
        Ra = mod1.HTP.GetRows
        mod1.HTP.Close
        Set mod1.HTP = Nothing
        On Error Resume Next
        La = UBound(Ra, 2)
        DHB.dtgDHB.Rows = La + 30
        DHB.dtgLL.Rows = DHB.dtgDHB.Rows
        DHB.dtgLL.Cols = DHB.dtgDHB.Cols
        
        For oo = 1 To La + 1
            DHB.dtgDHB.Row = oo: DHB.dtgLL.Row = oo
            For ii = 1 To 10
                DHB.dtgDHB.Col = ii: DHB.dtgLL.Col = ii
                DHB.dtgDHB.Text = Ra(ii - 1, oo - 1)
                DHB.dtgLL.Text = DHB.dtgDHB.Text
                If ii = 1 Then
                    If OBm <> DHB.dtgDHB.Text Then
                        OBm = DHB.dtgDHB.Text
                    Else
                        DHB.dtgDHB.Text = ""
                    End If
                End If
            Next
        Next

End Sub
