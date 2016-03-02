VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmLingjian 
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12195
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   12195
   Begin VB.CommandButton cmdKq 
      Caption         =   "修改"
      Height          =   285
      Left            =   11400
      TabIndex        =   9
      Top             =   7830
      Width           =   735
   End
   Begin VB.Frame frmMod 
      Caption         =   "修改"
      Height          =   1575
      Left            =   60
      TabIndex        =   7
      Top             =   6180
      Visible         =   0   'False
      Width           =   12105
      Begin VB.CommandButton cmdAll 
         Caption         =   "全部"
         Height          =   345
         Left            =   9840
         TabIndex        =   29
         Top             =   1080
         Width           =   1035
      End
      Begin VB.CommandButton cmdGx 
         Caption         =   "更新该记录"
         Height          =   345
         Left            =   6750
         TabIndex        =   28
         Top             =   1110
         Width           =   1125
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "添加新记录"
         Height          =   345
         Left            =   7980
         TabIndex        =   27
         Top             =   1080
         Width           =   1125
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除"
         Height          =   285
         Left            =   11340
         TabIndex        =   26
         Top             =   180
         Width           =   675
      End
      Begin VB.TextBox txtBz 
         Height          =   795
         Left            =   6690
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   180
         Width           =   3105
      End
      Begin VB.TextBox txtDj 
         Height          =   270
         Left            =   4530
         TabIndex        =   23
         Top             =   1170
         Width           =   1215
      End
      Begin VB.TextBox txtJJ 
         Height          =   270
         Left            =   4530
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtCj 
         Height          =   270
         Left            =   4530
         TabIndex        =   21
         Top             =   510
         Width           =   1215
      End
      Begin VB.TextBox txtBj 
         Height          =   270
         Left            =   4530
         TabIndex        =   20
         Top             =   180
         Width           =   1215
      End
      Begin VB.TextBox txtJzxh 
         Height          =   285
         Left            =   1380
         TabIndex        =   19
         Top             =   1080
         Width           =   1665
      End
      Begin VB.TextBox txtLjmc 
         Height          =   285
         Left            =   1380
         TabIndex        =   13
         Top             =   645
         Width           =   1665
      End
      Begin VB.TextBox txtLjbh 
         Height          =   285
         Left            =   1380
         TabIndex        =   11
         Top             =   210
         Width           =   1665
      End
      Begin VB.CommandButton cmdGB 
         Caption         =   "关闭"
         Height          =   315
         Left            =   11310
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "备注:"
         Height          =   225
         Left            =   6090
         TabIndex        =   24
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label10 
         Caption         =   "机组型号"
         Height          =   225
         Left            =   450
         TabIndex        =   18
         Top             =   1110
         Width           =   885
      End
      Begin VB.Label Label9 
         Caption         =   "底价"
         Height          =   285
         Left            =   3660
         TabIndex        =   17
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "进价"
         Height          =   225
         Left            =   3660
         TabIndex        =   16
         Top             =   870
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   "最低售价"
         Height          =   225
         Left            =   3660
         TabIndex        =   15
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "面价"
         Height          =   195
         Left            =   3660
         TabIndex        =   14
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "零件名称"
         Height          =   225
         Left            =   450
         TabIndex        =   12
         Top             =   690
         Width           =   885
      End
      Begin VB.Label Label3 
         Caption         =   "零件编号"
         Height          =   255
         Left            =   450
         TabIndex        =   10
         Top             =   270
         Width           =   1035
      End
   End
   Begin MSDataGridLib.DataGrid dtgView 
      Bindings        =   "frmLpg.frx":0000
      Height          =   6165
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   10874
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "LjBH"
         Caption         =   "零件编号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "LjMc"
         Caption         =   "零件名称"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Bj"
         Caption         =   "面价"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Cj"
         Caption         =   "最低售价"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "JJ"
         Caption         =   "进价"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "dj"
         Caption         =   "伏斯价"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Bz"
         Caption         =   "备注"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "XH"
         Caption         =   "机组型号"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Cou"
         Caption         =   "数量"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Lid"
         Caption         =   "Lid"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "xjf"
         Caption         =   "确认价"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "√"
            FalseValue      =   ""
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   7
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column10 
            Button          =   -1  'True
            ColumnWidth     =   599.811
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox comLx 
      Height          =   300
      ItemData        =   "frmLpg.frx":0015
      Left            =   1650
      List            =   "frmLpg.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   7800
      Width           =   1875
   End
   Begin VB.CommandButton cmdReq 
      Caption         =   "查  询"
      Height          =   315
      Left            =   6750
      TabIndex        =   1
      Top             =   7800
      Width           =   1125
   End
   Begin VB.ComboBox comJzXh 
      Height          =   300
      Left            =   4020
      TabIndex        =   0
      Top             =   7800
      Width           =   1905
   End
   Begin MSAdodcLib.Adodc adoLpg 
      Height          =   765
      Left            =   7830
      Top             =   7410
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1349
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\work.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\work\demo\work.mdb;Persist Security Info=False"
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
   Begin VB.Label Label2 
      Caption         =   "请业务员在询价时,一定要认清该零配件所属的机组型号"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   6930
      Width           =   4845
   End
   Begin VB.Label Label1 
      Caption         =   "值:"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   7860
      Width           =   405
   End
   Begin VB.Label Label4 
      Caption         =   "查询方式:"
      Height          =   195
      Left            =   750
      TabIndex        =   2
      Top             =   7860
      Width           =   1125
   End
End
Attribute VB_Name = "frmLingjian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LpXh As Object
Dim adoTT As String

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_Click()

End Sub

Private Sub cmdAdd_Click()
Dim liD As Double
Dim tt As String
Dim ii As Integer
On Error Resume Next
If txtLjbh.Text = "" Then txtLjbh.Text = "不详"
If txtLjmc.Text = "" Then txtLjmc.Text = "不详"
If txtJzxh.Text = "" Then txtJzxh.Text = "不详"
ii = MsgBox("是否添加此条新记录?", vbYesNo, "询问")
If ii = vbYes Then
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "LPGJia"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@jzpb") = frmLingjian.Caption
    mod1.cmd.Parameters("@jzxh") = Trim(txtJzxh.Text)
    mod1.cmd.Parameters("@ljbh") = Trim(txtLjbh.Text)
    mod1.cmd.Parameters("@ljmc") = Trim(txtLjmc.Text)
    mod1.cmd.Parameters("@bj") = Val(txtBj.Text)
    mod1.cmd.Parameters("@cj") = Val(txtCj.Text)
    mod1.cmd.Parameters("@jj") = Val(txtJJ.Text)
    mod1.cmd.Parameters("@dj") = Val(txtDj.Text)
    mod1.cmd.Parameters("@bz") = Trim(txtBz.Text)
    mod1.cmd.Parameters("@pjf") = 0
    mod1.cmd.Execute
    liD = mod1.cmd.Parameters("@lid").Value
    Set cmd = Nothing
    
    tt = "LpgOpen(" & liD & ")"
    adoLpg.Recordset.Close
    adoLpg.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    Set dtgView.DataSource = adoLpg

    adoTT = tt
End If
End Sub

Private Sub cmdAll_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
tt = "select * from LPG_KLV where jzpb='" & frmLingjian.Caption & "'"
adoLpg.Recordset.Close
adoLpg.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
Set dtgView.DataSource = adoLpg
End Sub

Private Sub cmdDel_Click()
Dim tt As String
Dim ii As Integer
On Error Resume Next
ii = MsgBox("是否确认删除此记录?", vbYesNo + vbInformation, "询问")
If ii = vbYes Then
    tt = "delete from lpg where lid=" & adoLpg.Recordset.Fields("lid").Value
    Set mod1.HTP = CreateObject("adodb.recordset")
    mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
    '更新列表
        adoLpg.Recordset.Close
        adoLpg.Recordset.Open adoTT, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
        Set dtgView.DataSource = adoLpg
    dtgView.Refresh
End If
End Sub

Private Sub cmdGB_Click()
frmMod.Visible = False
cmdKq.Visible = True
End Sub

Private Sub cmdGx_Click()
Dim tt As String
On Error Resume Next
If txtLjbh.Text = "" Then Exit Sub
'''If adoLpg.Recordset.Fields("xjf").Value = True And Val(txtJJ.Text) > txtJJ.Tag Then
'''    MsgBox ("此确认价格大于以前的价格,请与宋经理联系修改!")
'''    Exit Sub
'''End If
    Set mod1.cmd = CreateObject("adodb.command")
    mod1.cmd.ActiveConnection = mod1.cc
    mod1.cmd.CommandText = "LPGGx"
    mod1.cmd.CommandType = adCmdStoredProc
    mod1.cmd.Parameters("@jzpb") = frmLingjian.Caption
    mod1.cmd.Parameters("@jzxh") = Trim(txtJzxh.Text)
    mod1.cmd.Parameters("@ljbh") = Trim(txtLjbh.Text)
    mod1.cmd.Parameters("@ljmc") = Trim(txtLjmc.Text)
    mod1.cmd.Parameters("@bj") = Val(txtBj.Text)
    mod1.cmd.Parameters("@cj") = Val(txtCj.Text)
    mod1.cmd.Parameters("@jj") = Val(txtJJ.Text)
    mod1.cmd.Parameters("@dj") = Val(txtDj.Text)
    mod1.cmd.Parameters("@bz") = Trim(txtBz.Text)
    mod1.cmd.Parameters("@pjf") = 0
    mod1.cmd.Parameters("@lid") = adoLpg.Recordset.Fields("lid").Value
    mod1.cmd.Execute
    Set cmd = Nothing
    
    tt = "LpgOpen(" & adoLpg.Recordset.Fields("lid").Value & ")"
    adoLpg.Recordset.Close
    adoLpg.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    Set dtgView.DataSource = adoLpg
    
    txtJzxh.Text = ""
    txtLjbh.Text = ""
    txtLjmc.Text = ""
    txtBj.Text = ""
    txtCj.Text = ""
    txtJJ.Text = ""
    txtDj.Text = ""
    txtBz.Text = ""
cmdGx.Enabled = False
End Sub

Private Sub cmdKq_Click()
frmMod.Visible = True
cmdKq.Visible = False
End Sub

Private Sub cmdReq_Click()
Dim tt As String
On Error Resume Next
Select Case comLx.Text
    Case "机组型号"
        tt = "LPG_KLV_xh('" & comJzXh.Text & "','" & frmLingjian.Caption & "')"
    Case "零配件编号"
        tt = "LPG_KLV_ljbh('" & comJzXh.Text & "','" & frmLingjian.Caption & "')"
    Case "零配件名称"
        tt = "LPG_KLV_ljmc('" & comJzXh.Text & "','" & frmLingjian.Caption & "')"
End Select
    adoLpg.Recordset.Close
    adoLpg.Recordset.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdStoredProc
    Set dtgView.DataSource = adoLpg
    adoTT = tt
End Sub

Private Sub comLx_Click()
Dim oo As Integer
On Error Resume Next
For oo = comJzXh.ListCount - 1 To 0 Step -1
    comJzXh.RemoveItem oo
Next
    
If comLx.Text = "机组型号" Then

    LpXh.MoveFirst
    Do While Not LpXh.EOF
        comJzXh.AddItem LpXh.Fields("jzXh").Value
        LpXh.MoveNext
    Loop
    If frmLingjian.Caption = "开利" Then
        comJzXh.RemoveItem 3
        comJzXh.AddItem "30HK"
        comJzXh.AddItem "30HR"
    End If
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub dtgView_ButtonClick(ByVal ColIndex As Integer)
Dim tt As String
On Error Resume Next
If mod1.VLP <> 3 Then
    Exit Sub
End If
If adoLpg.Recordset.Fields("xjf").Value = True Then
    tt = "update lpg set xjf=0 where lid=" & adoLpg.Recordset.Fields("lid").Value
Else
    tt = "update lpg set xjf=1 where lid=" & adoLpg.Recordset.Fields("lid").Value
End If
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
adoLpg.Recordset.Requery
Set dtgView.DataSource = adoLpg
End Sub

Private Sub dtgView_Click()
On Error Resume Next
If IsNull(adoLpg.Recordset.Fields("lid").Value) = False Then
    txtJzxh.Text = ""
    txtLjbh.Text = ""
    txtLjmc.Text = ""
    txtBj.Text = ""
    txtCj.Text = ""
    txtJJ.Text = ""
    txtDj.Text = ""
    txtBz.Text = ""

    txtJzxh.Text = adoLpg.Recordset.Fields("xh").Value
    txtLjbh.Text = adoLpg.Recordset.Fields("ljbh").Value
    txtLjmc.Text = adoLpg.Recordset.Fields("ljmc").Value
    txtBj.Text = adoLpg.Recordset.Fields("bj").Value
    txtCj.Text = adoLpg.Recordset.Fields("cj").Value
    txtJJ.Text = adoLpg.Recordset.Fields("jj").Value
    txtDj.Text = adoLpg.Recordset.Fields("dj").Value
    txtJJ.Tag = adoLpg.Recordset.Fields("jj").Value
    txtBz.Text = adoLpg.Recordset.Fields("bz").Value
    
    cmdGx.Enabled = True
Else
    cmdGx.Enabled = False
End If
End Sub

Private Sub dtgView_DblClick()
If frmGXBj.Visible = True Then
    frmGXBj.comJzXh.Text = adoLpg.Recordset.Fields("xh").Value
    frmGXBj.txtLjmc.Text = adoLpg.Recordset.Fields("ljmc").Value
    frmGXBj.txtLjbh.Text = adoLpg.Recordset.Fields("ljbh").Value
    If adoLpg.Recordset.Fields("xjf").Value = True Then
        frmGXBj.txtMj.Text = adoLpg.Recordset.Fields("bj").Value
        frmGXBj.txtDj.Text = adoLpg.Recordset.Fields("jj").Value
    End If
    frmGXBj.frmCg.Enabled = True
    frmGXBj.comLx.Text = "零配件"
End If
    
End Sub

Private Sub dtgView_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
If IsNull(adoLpg.Recordset.Fields("lid").Value) = False Then
    txtJzxh.Text = ""
    txtLjbh.Text = ""
    txtLjmc.Text = ""
    txtBj.Text = ""
    txtCj.Text = ""
    txtJJ.Text = ""
    txtDj.Text = ""
    txtBz.Text = ""

    txtJzxh.Text = adoLpg.Recordset.Fields("xh").Value
    txtLjbh.Text = adoLpg.Recordset.Fields("ljbh").Value
    txtLjmc.Text = adoLpg.Recordset.Fields("ljmc").Value
    txtBj.Text = adoLpg.Recordset.Fields("bj").Value
    txtCj.Text = adoLpg.Recordset.Fields("cj").Value
    txtJJ.Text = adoLpg.Recordset.Fields("jj").Value
    txtDj.Text = adoLpg.Recordset.Fields("dj").Value
    txtJJ.Tag = adoLpg.Recordset.Fields("jj").Value
    txtBz.Text = adoLpg.Recordset.Fields("bz").Value
    
    cmdGx.Enabled = True
Else
    cmdGx.Enabled = False
End If
End Sub


Private Sub Form_Load()

frmLingjian.Width = 12315
frmLingjian.Height = 8625

End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
frmLingjian.Visible = False
frmZu.Enabled = True
Cancel = True
End If
End Sub

