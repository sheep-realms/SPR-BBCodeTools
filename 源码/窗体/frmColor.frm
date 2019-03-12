VERSION 5.00
Begin VB.Form frmColor 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "调色板"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdBCSave 
      Caption         =   "保存"
      Height          =   495
      Left            =   2760
      TabIndex        =   16
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "设为前景色"
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      Caption         =   "注意事项"
      Height          =   2535
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   2655
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "由于DZ编辑器所见机所得模式下无法渲染RGB函数，因此无法预览效果。"
         Height          =   2175
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdCSave 
      Caption         =   "保存"
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdBC 
      Caption         =   "设为背景色"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   2655
   End
   Begin VB.PictureBox pF 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3600
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   9
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox p0 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2880
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox pColor 
      Height          =   375
      Left            =   5280
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   6
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox pLast 
      Height          =   375
      Left            =   4560
      ScaleHeight     =   315
      ScaleWidth      =   675
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox tC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6120
      TabIndex        =   3
      Text            =   "0"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox tC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   2
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox tC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   6120
      TabIndex        =   1
      Text            =   "0"
      Top             =   240
      Width           =   615
   End
   Begin VB.PictureBox pC 
      AutoRedraw      =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   2880
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   255
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   3135
   End
   Begin VB.PictureBox pC 
      AutoRedraw      =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   2880
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   255
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   3135
   End
   Begin VB.PictureBox pC 
      AutoRedraw      =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2880
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   255
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   3135
   End
   Begin VB.Menu m0 
      Caption         =   "取色器"
   End
   Begin VB.Menu m1 
      Caption         =   "RGB"
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim r As Integer, g As Integer, b As Integer
Dim clk(2) As Boolean

Private Function SetRGB()
On Error Resume Next
    If IsNumeric(r) = False Then r = 0
    If IsNumeric(g) = False Then g = 0
    If IsNumeric(b) = False Then b = 0
    If r < 0 Then r = 0
    If g < 0 Then g = 0
    If b < 0 Then b = 0
    If r > 255 Then r = 255
    If g > 255 Then g = 255
    If b > 255 Then b = 255
    Dim i As Integer
    For i = 0 To 255
        If r = i Then pC(0).Line (i - 1, 0)-(i + 1, 1), RGB(255, 255, 255), B Else pC(0).Line (i, 0)-(i, 1), RGB(i, g, b), B
        If g = i Then pC(1).Line (i - 1, 0)-(i + 1, 1), RGB(255, 255, 255), B Else pC(1).Line (i, 0)-(i, 1), RGB(r, i, b), B
        If b = i Then pC(2).Line (i - 1, 0)-(i + 1, 1), RGB(255, 255, 255), B Else pC(2).Line (i, 0)-(i, 1), RGB(r, g, i), B
        'pC(0).Line (i, 0)-(i, 1), RGB(i, g, b), B
        'pC(1).Line (i, 0)-(i, 1), RGB(r, i, b), B
        'pC(2).Line (i, 0)-(i, 1), RGB(r, g, i), B
    Next i
    SetColor
End Function

Private Function SetColor()
    pColor.BackColor = RGB(r, g, b)
    txtInput.Text = "RGB(" & r & "," & g & "," & b & ")"
    tC(0).Text = r
    tC(1).Text = g
    tC(2).Text = b
End Function

Private Sub cmdCSave_Click()
    TextColor(0) = r
    TextColor(1) = g
    TextColor(2) = b
    frm.txtMsg.SetFocus
    ColorReset
    Unload Me
End Sub

Private Sub cmdColor_Click()
    TextColor(0) = r
    TextColor(1) = g
    TextColor(2) = b
    BBCode "color", , "RGB(" & r & "," & g & "," & b & ")"
    ColorReset
    Unload Me
End Sub

Private Sub Form_Load()
    r = TextColor(0)
    g = TextColor(1)
    b = TextColor(2)
    
    SetRGB
    txtInput.Text = "RGB(" & r & "," & g & "," & b & ")"
    pLast.BackColor = RGB(TextColor(0), TextColor(1), TextColor(2))
    If frm.chkTop = 1 Then TopMode = True: frm.chkTop.Value = 0
    frm.chkTop.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If TopMode = True Then TopMode = False: frm.chkTop.Value = 1
    frm.chkTop.Enabled = True
End Sub

Private Sub p0_Click()
    r = 0
    g = 0
    b = 0
    SetRGB
End Sub

Private Sub pC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    clk(Index) = True
End Sub

Private Sub pC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If clk(Index) = True Then
        Dim z As Integer
        z = Int(X)
        If IsNumeric(z) = False Then z = 0
        If z > 255 Then z = 255
        If z < 0 Then z = 0
        tC(Index).Text = z
    End If
End Sub



Private Sub pC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    clk(Index) = False
    If Index = 0 Then r = Int(X)
    If Index = 1 Then g = Int(X)
    If Index = 2 Then b = Int(X)
    SetRGB
End Sub

Private Sub pF_Click()
    r = 255
    g = 255
    b = 255
    SetRGB
End Sub

Private Sub pLast_Click()
    r = TextColor(0)
    g = TextColor(1)
    b = TextColor(2)
    SetRGB
End Sub

Private Sub tC_LostFocus(Index As Integer)
    Dim z As Integer
    z = tC(Index).Text
    If IsNumeric(z) = False Then z = 0
    If z > 255 Then z = 255
    If z < 0 Then z = 0
    If Index = 0 Then r = z
    If Index = 1 Then g = z
    If Index = 2 Then b = z
    SetRGB
End Sub
