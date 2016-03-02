VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmPeiBB1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "培训课程统计表"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   15210
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      Height          =   765
      Left            =   14550
      Picture         =   "frmPeiBB1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   585
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dtgBr 
      Height          =   8235
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   14526
      _Version        =   393216
      BackColor       =   12648384
      Rows            =   30
      Cols            =   18
      FixedCols       =   2
      BackColorFixed  =   16777152
      BackColorBkg    =   12648384
      WordWrap        =   -1  'True
      SelectionMode   =   1
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   18
   End
End
Attribute VB_Name = "frmPeiBB1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub dtgBRFF()
dtgBr.Clear
dtgBr.Rows = 50
dtgBr.Row = 0
dtgBr.Col = 0: dtgBr.Text = "序号": dtgBr.CellFontBold = True
dtgBr.Col = 1: dtgBr.Text = "培训类别": dtgBr.CellFontBold = True
dtgBr.Col = 2: dtgBr.Text = "课程名称": dtgBr.CellFontBold = True
dtgBr.Col = 3: dtgBr.Text = "培训时间": dtgBr.CellFontBold = True
dtgBr.Col = 4: dtgBr.Text = "培训地点": dtgBr.CellFontBold = True
dtgBr.Col = 5: dtgBr.Text = "课程时数": dtgBr.CellFontBold = True
dtgBr.Col = 6: dtgBr.Text = "主办单位": dtgBr.CellFontBold = True
dtgBr.Col = 7: dtgBr.Text = "讲师": dtgBr.CellFontBold = True
dtgBr.Col = 8: dtgBr.Text = "学员对象": dtgBr.CellFontBold = True
dtgBr.Col = 9: dtgBr.Text = "参加人数": dtgBr.CellFontBold = True
dtgBr.Col = 10: dtgBr.Text = "实到人数": dtgBr.CellFontBold = True
dtgBr.Col = 11: dtgBr.Text = "出勤率": dtgBr.CellFontBold = True
dtgBr.Col = 12: dtgBr.Text = "培训费用": dtgBr.CellFontBold = True
dtgBr.Col = 13: dtgBr.Text = "人均费用": dtgBr.CellFontBold = True
dtgBr.Col = 14: dtgBr.Text = "培训满意度": dtgBr.CellFontBold = True
dtgBr.Col = 15: dtgBr.Text = "培训评估": dtgBr.CellFontBold = True
dtgBr.Col = 16: dtgBr.Text = "总时数": dtgBr.CellFontBold = True
dtgBr.Col = 17: dtgBr.Text = "备注": dtgBr.CellFontBold = True
End Sub

Public Sub Bound(Ra)
Dim La As Long
Dim oo As Long
Dim ii As Integer
Dim Pid As Long
Dim XYuan As String
Dim Cs As Integer
Dim SS As Integer
Dim Line As Long
On Error Resume Next
Pid = 0
Call dtgBRFF
dtgBr.Visible = False
La = UBound(Ra, 2) + 1
dtgBr.Rows = La + 10
Line = 0
For oo = 1 To La
    If Pid <> Ra(18, oo - 1) Then
     Line = Line + 1:   dtgBr.Row = Line
     dtgBr.RowHeight(oo) = dtgBr.RowHeight(0) * 3
        For ii = 0 To 17
            dtgBr.Col = ii
            dtgBr.Text = Ra(ii, oo - 1)
        Next
        XYuan = Ra(8, oo - 1): dtgBr.Col = 8: dtgBr.Text = XYuan
        dtgBr.Col = 9: dtgBr.Text = 1 '参加人数置1
        Cs = 1
        Pid = Ra(18, oo - 1)
        If Ra(19, oo - 1) = True Then ' 实到人数
            dtgBr.Col = 10: dtgBr.Text = 1: SS = 1
        Else
            dtgBr.Col = 10: dtgBr.Text = 0: SS = 0
        End If
        dtgBr.Col = 11: dtgBr.Text = Str(Round(SS / Cs, 2) * 100) & "%"
    Else

        XYuan = XYuan & " " & Ra(8, oo - 1): dtgBr.Col = 8: dtgBr.Text = XYuan
        Cs = Cs + 1: dtgBr.Col = 9: dtgBr.Text = Cs
        If Ra(19, oo - 1) = True Then ' 实到人数
            SS = SS + 1: dtgBr.Col = 10: dtgBr.Text = SS
        End If
        dtgBr.Col = 11: dtgBr.Text = Str(Round(SS / Cs, 2) * 100) & "%"
    End If
Next



dtgBr.Visible = True
End Sub

Private Sub cmdBack_Click()
Me.Visible = False
End Sub

Private Sub Form_Load()
Me.Width = mod1.FWidth
Me.Height = mod1.FHeight
Me.Left = 0
Me.Top = 0
End Sub


