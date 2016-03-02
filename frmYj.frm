VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmYj 
   Caption         =   "佣金表"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2610
   ScaleWidth      =   7515
   Begin VB.CommandButton cmdSave 
      Height          =   405
      Left            =   7080
      Picture         =   "frmYj.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2220
      Visible         =   0   'False
      Width           =   435
   End
   Begin MSAdodcLib.Adodc adoYj 
      Height          =   330
      Left            =   4050
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
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
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除"
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "添加"
      Height          =   315
      Left            =   810
      TabIndex        =   1
      Top             =   2280
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid dtgYj 
      Height          =   2235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   3942
      _Version        =   393216
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "yED"
         Caption         =   "收到款额度"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   5
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "YingFu"
         Caption         =   "相应支付佣金"
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
         DataField       =   "FF"
         Caption         =   "支付否"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "OK!"
            FalseValue      =   "未付"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   7
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "zFu"
         Caption         =   "实际支付"
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
         DataField       =   "fRQ"
         Caption         =   "支付日期"
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
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         BeginProperty Column00 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataField       =   "UserId"
      DataSource      =   "adoYj"
      Height          =   315
      Left            =   150
      TabIndex        =   5
      Top             =   2340
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblKhmc 
      Caption         =   "Label2"
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   2310
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblHtbh 
      Caption         =   "Label1"
      Height          =   225
      Left            =   2400
      TabIndex        =   3
      Top             =   2340
      Visible         =   0   'False
      Width           =   1125
   End
End
Attribute VB_Name = "frmYj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

On Error Resume Next
If mod1.DName = "马晓聪" Then '由我来添加已经支付过的奖金明细
    adoYj.Recordset.AddNew "htbh", lblhtbh.Caption
    adoYj.Recordset.Update "khmc", lblKhmc.Caption
    adoYj.Recordset.Update "FF", 1
    adoYj.Recordset.Update "pwf", 1
    adoYj.Recordset.Update "frq", "2006-1-1"
    'adoYj.Recordset.UpdateBatch
    Set dtgYJ.DataSource = adoYj
ElseIf mod1.DName = "倪旭" Then
    
    adoYj.Recordset.AddNew "htbh", lblhtbh.Caption
    adoYj.Recordset.Update "khmc", lblKhmc.Caption
    adoYj.Recordset.Update "FF", 0
    'adoYj.Recordset.UpdateBatch
    Set dtgYJ.DataSource = adoYj
End If

End Sub

Private Sub cmdDel_Click()
On Error Resume Next
If adoYj.Recordset.Fields("pwf").Value = True Then
    MsgBox ("此笔金额已经支付,不能删除!")
    Exit Sub
End If
adoYj.Recordset.Delete adAffectCurrent
End Sub

Private Sub cmdSave_Click()
Dim Yj As Long
Dim LR As Long
Dim tt As String
On Error Resume Next
tt = "select yjff from htping where htbh='" & lblhtbh.Caption & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockReadOnly, adCmdText
If IsNull(mod1.HTP.RecordCount) Or mod1.HTP.RecordCount = 0 Then
    Exit Sub
End If
If mod1.HTP.Fields("yjff").Value = True Then
    MsgBox ("奖金已经全部支付,不能再更改!")
    Exit Sub
End If
Yj = 0
adoYj.Recordset.UpdateBatch
adoYj.Recordset.MoveFirst
Do While Not adoYj.Recordset.EOF
    Yj = Yj + adoYj.Recordset.Fields("YingFu").Value
    adoYj.Recordset.MoveNext
Loop
If form2Htp.Visible = True Then
    form2Htp.txtYj1.Text = Yj
    form2Htp.txtLr1.Text = Val(form2Htp.txtJlr1.Text) - Yj
    LR = form2Htp.txtLr1.Text
ElseIf wbHTP.Visible = True Then
    wbHTP.txtYj1.Text = Yj
    wbHTP.txtLr1.Text = Val(wbHTP.txtJlr1.Text) - Yj
    LR = wbHTP.txtLr1.Text
End If
adoYj.Recordset.MoveFirst
tt = "update htping set yj=" & Yj & ",xmLr=" & LR & " where htbh='" & adoYj.Recordset.Fields("htbh").Value & "'"
Set mod1.HTP = CreateObject("adodb.recordset")
mod1.HTP.Open tt, mod1.workKK, adOpenKeyset, adLockBatchOptimistic, adCmdText
MsgBox "ok!"
End Sub

Private Sub dtgYj_AfterColEdit(ByVal ColIndex As Integer)
On Error Resume Next
If ColIndex = 0 Then
adoYj.Recordset.Fields("yED").Value = adoYj.Recordset.Fields("yED").Value / 100
End If

End Sub

Private Sub dtgYj_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If adoYj.Recordset.Fields("pwf").Value = True Or adoYj.Recordset.Fields("ff").Value = True Then
    Cancel = True
End If
End Sub

Private Sub dtgYj_BeforeDelete(Cancel As Integer)
On Error Resume Next
If adoYj.Recordset.Fields("pwf").Value = True Or adoYj.Recordset.Fields("ff").Value = True Then
    Cancel = True
End If
End Sub


Private Sub dtgYj_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub


Private Sub Form_Load()
frmYj.Height = 3015
frmYj.Width = 7635
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
frmYj.Visible = False
Cancel = True
End If
End Sub
