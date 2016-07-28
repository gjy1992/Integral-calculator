VERSION 5.00
Begin VB.Form 积分计算器 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "积分计算器"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   8
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "复制结果"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   7
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   1680
      Width           =   6495
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2760
      Width           =   6495
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   3840
      Width           =   6495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   5760
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "表达式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "积分下限"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "积分上限"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "请用大写字母X代表积分变量,积分下限和积分上限都可以是不带未知量的表达式"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "积分计算器"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const n = 5
Dim t(1 To 5) As Double, x(1 To 5) As Double

Private Sub Command1_Click()
Dim a As Double, b As Double, S As Double, SS As Double, Serr As String
Call Calc(Text2.Text, a, Serr)
If Serr <> "" Then Label4.Caption = "错误": Exit Sub
Call Calc(Text3.Text, b, Serr)
If Serr <> "" Then Label4.Caption = "错误": Exit Sub
For i = 1 To 5
Call Calc(Text1.Text, SS, Serr, (a + b) / 2 + (b - a) / 2 * x(i))
If Serr = "" Then S = S + t(i) * SS Else Label4.Caption = "错误": Exit Sub
Next i
S = S * (b - a) / 2
Label4.Caption = Trim(Str(S))
End Sub

Private Sub Command2_Click()
Clipboard.SetText (Label4.Caption)
End Sub

Private Sub Form_Load()
//高斯积分公式中的参数
t(1) = 0.2369269
t(2) = 0.4786287
t(3) = 0.5688888
t(4) = 0.4786287
t(5) = 0.2369269
x(1) = -0.906179845938645
x(2) = -0.538469310105683
x(3) = 0
x(4) = 0.538469310105683
x(5) = 0.906179845938645
Show form1
Text1.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1.SetFocus: Command1_Click
End Sub

