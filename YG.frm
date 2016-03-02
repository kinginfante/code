VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EF977422-E047-42A7-A004-1C0695C81FCF}#1.0#0"; "NiceForm.ocx"
Begin VB.Form frmYg 
   Caption         =   "员工守则"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   15240
   Begin NiceFormControl.NiceForm NF 
      Left            =   14430
      Top             =   3090
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin NiceFormControl.NiceOption NiceOption3 
      Height          =   240
      Left            =   13770
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "财务制度"
      SkinIdx         =   5
   End
   Begin NiceFormControl.NiceOption NiceOption2 
      Height          =   240
      Left            =   13770
      TabIndex        =   3
      Top             =   1560
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "人事制度"
      SkinIdx         =   5
   End
   Begin NiceFormControl.NiceOption NiceOption1 
      Height          =   240
      Left            =   13770
      TabIndex        =   2
      Top             =   930
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "行政制度"
      SkinIdx         =   5
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "返回"
      Height          =   585
      Left            =   14490
      Picture         =   "YG.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8520
      Width           =   675
   End
   Begin RichTextLib.RichTextBox riha 
      Height          =   9165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   16166
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"YG.frx":0102
   End
   Begin MSAdodcLib.Adodc adoFile 
      Height          =   330
      Left            =   13500
      Top             =   330
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Left            =   14340
      TabIndex        =   5
      Top             =   4170
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "frmYg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
frmYg.Visible = False
frmZu.TBa.Buttons(6).Value = tbrUnpressed
End Sub

Private Sub Form_Load()
NF.LoadSkin 5
Me.Height = mod1.FHeight + 100
Me.Width = mod1.FWidth
frmYg.Left = 0
frmYg.Top = 0
frmYg.Width = mod1.FWidth
frmYg.Height = mod1.FHeight


'riha.LoadFile App.Path & "\j1.mdb", 0
Call NiceOption1_Click
Me.NiceOption1.Value = True
Me.Height = Me.Height + 500
Me.Width = Me.Width + 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MDI.Cq = False Then
frmYg.Visible = False
frmZu.TBa.Buttons(6).Value = tbrUnpressed
Cancel = True
End If
End Sub

Private Sub NiceOption1_Click()
Dim bt() As Byte
Dim tt As String
On Error Resume Next




'如果不存在,则下载文件
 If Dir(App.Path & "\j1.mdb") = "" Then
 
 tt = "select Nfile,fsize from HMFile where Fname='j1.mdb' "
adoFile.Recordset.Close
adoFile.Recordset.Open tt, mod1.wzcc, adOpenKeyset, adLockReadOnly, adCmdText
ReDim bt(adoFile.Recordset.Fields("Fsize").Value) As Byte
bt() = adoFile.Recordset.Fields("Nfile").GetChunk(adoFile.Recordset.Fields("Fsize").Value + 1)


Open (App.Path & "\j1.mdb") For Binary As #5
Put #5, , bt()
Close #5
 

End If
riha.LoadFile App.Path & "\j1.mdb", 0
End Sub

Private Sub NiceOption2_Click()
Dim bt() As Byte
Dim tt As String
On Error Resume Next




'如果不存在,则下载文件
 If Dir(App.Path & "\j2.mdb") = "" Then
 
 tt = "select Nfile,fsize from HMFile where Fname='j2.mdb' "
adoFile.Recordset.Close
adoFile.Recordset.Open tt, mod1.wzcc, adOpenKeyset, adLockReadOnly, adCmdText
ReDim bt(adoFile.Recordset.Fields("Fsize").Value) As Byte
bt() = adoFile.Recordset.Fields("Nfile").GetChunk(adoFile.Recordset.Fields("Fsize").Value + 1)


Open (App.Path & "\j2.mdb") For Binary As #5
Put #5, , bt()
Close #5
 

End If
riha.LoadFile App.Path & "\j2.mdb", 0
End Sub

Private Sub NiceOption3_Click()
Dim bt() As Byte
Dim tt As String
On Error Resume Next




'如果不存在,则下载文件
 If Dir(App.Path & "\j3.mdb") = "" Then
 
 tt = "select Nfile,fsize from HMFile where Fname='j3.mdb' "
adoFile.Recordset.Close
adoFile.Recordset.Open tt, mod1.wzcc, adOpenKeyset, adLockReadOnly, adCmdText
ReDim bt(adoFile.Recordset.Fields("Fsize").Value) As Byte
bt() = adoFile.Recordset.Fields("Nfile").GetChunk(adoFile.Recordset.Fields("Fsize").Value + 1)


Open (App.Path & "\j3.mdb") For Binary As #5
Put #5, , bt()
Close #5
 

End If
riha.LoadFile App.Path & "\j3.mdb", 0
End Sub

Private Sub riha_Click()
cmdBack.SetFocus
End Sub

Private Sub riha_KeyDown(KeyCode As Integer, Shift As Integer)
cmdBack.SetFocus
End Sub
