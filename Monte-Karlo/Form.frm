VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Монте-Карло"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   5175
      Left            =   240
      TabIndex        =   32
      Top             =   240
      Width           =   2895
      Begin VB.Frame Frame2 
         Caption         =   "Границы"
         Height          =   3375
         Left            =   360
         TabIndex        =   36
         Top             =   360
         Width           =   1335
         Begin VB.TextBox Text1 
            Height          =   615
            Left            =   120
            TabIndex        =   38
            Text            =   "b"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Height          =   615
            Left            =   120
            TabIndex        =   37
            Text            =   "a"
            Top             =   2640
            Width           =   975
         End
         Begin VB.Image Image1 
            Height          =   1545
            Index           =   0
            Left            =   360
            Picture         =   "Form.frx":0000
            Top             =   960
            Width           =   450
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   855
         Left            =   1800
         TabIndex        =   35
         Top             =   3960
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Caption         =   "Точки"
         Height          =   975
         Left            =   360
         TabIndex        =   33
         Top             =   3840
         Width           =   1335
         Begin VB.TextBox Text9 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   34
            Text            =   "Form.frx":0450
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Решение"
      Height          =   1815
      Left            =   3480
      TabIndex        =   30
      Top             =   3600
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox Text6 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   31
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ввод функции"
      Height          =   3135
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   2880
         TabIndex        =   40
         Text            =   "1"
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Готово"
         Height          =   495
         Left            =   3480
         TabIndex        =   22
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Text            =   "1"
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   960
         TabIndex        =   20
         Text            =   "1"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2880
         TabIndex        =   19
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Text            =   "0"
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   4680
         TabIndex        =   17
         Text            =   "1"
         Top             =   1560
         Width           =   615
      End
      Begin VB.Frame Frame4 
         Height          =   2775
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1095
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   255
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Option3"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   255
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Option4"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   255
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Option5"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1680
            Width           =   255
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Option6"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   2040
            Width           =   255
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Option7"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "1"
            Height          =   375
            Left            =   480
            TabIndex        =   16
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Sin"
            Height          =   375
            Left            =   480
            TabIndex        =   15
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Cos"
            Height          =   255
            Left            =   480
            TabIndex        =   14
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "Tg"
            Height          =   375
            Left            =   480
            TabIndex        =   13
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Ctg"
            Height          =   255
            Left            =   480
            TabIndex        =   12
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label Label7 
            Caption         =   "Ln"
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "Exp"
            Height          =   255
            Left            =   480
            TabIndex        =   10
            Top             =   2400
            Width           =   495
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Степень"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "("
         Height          =   255
         Left            =   840
         TabIndex        =   29
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "x"
         Height          =   375
         Left            =   1560
         TabIndex        =   28
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "^"
         Height          =   375
         Left            =   2760
         TabIndex        =   27
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "[(               X + "
         Height          =   375
         Left            =   2760
         TabIndex        =   26
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label20 
         Caption         =   ") ^"
         Height          =   255
         Left            =   4440
         TabIndex        =   25
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label21 
         Caption         =   "] )"
         Height          =   375
         Left            =   5280
         TabIndex        =   24
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "^"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   23
         Top             =   1560
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.Image Image1 
      Height          =   2400
      Index           =   1
      Left            =   9840
      Picture         =   "Form.frx":0470
      Top             =   360
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   $"Form.frx":1056
      Height          =   1455
      Index           =   1
      Left            =   9720
      TabIndex        =   39
      Top             =   3480
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, Sum, X, Y, Integ, Integ2, max, min, D, i As Double, N, K As Long, str As String, Func, Chk As Integer

Private Sub Check1_Click()
If Text3.Visible = False Then
    Text3.Visible = True
    Label1(0).Visible = True
Else
    Text3.Visible = False
    Label1(0).Visible = Flase
End If
End Sub

Private Sub Form_Load()
Chk = 0
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub

Private Sub Text4_Click()
Text4.Text = ""
End Sub

Private Sub Text5_Click()
Text5.Text = ""
End Sub

Private Sub Text7_Click()
Text7.Text = ""
End Sub

Private Sub Text8_Click()
Text8.Text = ""
End Sub

Private Sub Text9_Click()
If Chk = 0 Then
Text9.Text = ""
Chk = 1
End If
End Sub

Private Sub Text10_Click()
Text10.Text = ""
End Sub

Private Sub Command1_Click()
a = Val(Text2.Text)
b = Val(Text1.Text)
N = Val(Text9.Text)
Integ = 0
Integ2 = 0
Frame3.Visible = True
End Sub

Private Sub Command2_Click()
If Option1.Value = True Then Func = 1
If Option2.Value = True Then Func = 2
If Option3.Value = True Then Func = 3
If Option4.Value = True Then Func = 4
If Option5.Value = True Then Func = 5
If Option6.Value = True Then Func = 6
If Option7.Value = True Then Func = 7

If Check1.Value = 1 Then

Select Case Func
Case 1
    max = Val(Text3.Text) ^ ((Val(Text4.Text) * (((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = Val(Text3.Text) ^ ((Val(Text4.Text) * (((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = Val(Text3.Text) ^ ((Val(Text4.Text) * (((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= Val(Text3.Text) ^ ((Val(Text4.Text) * (((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N
    
Case 2
    max = Val(Text3.Text) ^ ((Val(Text4.Text) * (Sin((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (Sin((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = Val(Text3.Text) ^ ((Val(Text4.Text) * (Sin((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (Sin((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = Val(Text3.Text) ^ ((Val(Text4.Text) * (Sin((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= Val(Text3.Text) ^ ((Val(Text4.Text) * (Sin((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N
    
Case 3
    max = Val(Text3.Text) ^ ((Val(Text4.Text) * (Cos((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (Cos((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = Val(Text3.Text) ^ ((Val(Text4.Text) * (Cos((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (Cos((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = Val(Text3.Text) ^ ((Val(Text4.Text) * (Cos((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= Val(Text3.Text) ^ ((Val(Text4.Text) * (Cos((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N
    
Case 4
    max = Val(Text3.Text) ^ ((Val(Text4.Text) * (Tan((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = Val(Text3.Text) ^ ((Val(Text4.Text) * (Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = Val(Text3.Text) ^ ((Val(Text4.Text) * (Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= Val(Text3.Text) ^ ((Val(Text4.Text) * (Tan((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N
    
Case 5
    max = Val(Text3.Text) ^ ((Val(Text4.Text) * (Val(Text10.Text) / Tan((a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (Val(Text10.Text) / Tan((i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = Val(Text3.Text) ^ ((Val(Text4.Text) * (1 / Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (Val(Text10.Text) / Tan((i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = Val(Text3.Text) ^ ((Val(Text4.Text) * (1 / Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= Val(Text3.Text) ^ ((Val(Text4.Text) * (Val(Text10.Text) / Tan((X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N
    
Case 6
    max = Val(Text3.Text) ^ ((Val(Text4.Text) * (Log((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (Log((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = Val(Text3.Text) ^ ((Val(Text4.Text) * (Log((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (Log((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = Val(Text3.Text) ^ ((Val(Text4.Text) * (Log((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= Val(Text3.Text) ^ ((Val(Text4.Text) * (Log((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N
    
Case 7
    max = Val(Text3.Text) ^ ((Val(Text4.Text) * (Exp((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (Exp((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = Val(Text3.Text) ^ ((Val(Text4.Text) * (Exp((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If Val(Text3.Text) ^ ((Val(Text4.Text) * (Exp((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = Val(Text3.Text) ^ ((Val(Text4.Text) * (Exp((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= Val(Text3.Text) ^ ((Val(Text4.Text) * (Exp((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N

End Select

Else

Select Case Func
Case 1
    max = ((Val(Text4.Text) * (((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If ((Val(Text4.Text) * (((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = ((Val(Text4.Text) * (((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If ((Val(Text4.Text) * (((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = ((Val(Text4.Text) * (((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= ((Val(Text4.Text) * (((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N

Case 2
    max = ((Val(Text4.Text) * (Sin((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If ((Val(Text4.Text) * (Sin((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = ((Val(Text4.Text) * (Sin((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If ((Val(Text4.Text) * (Sin((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = ((Val(Text4.Text) * (Sin((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= ((Val(Text4.Text) * (Sin((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N
   
Case 3
    max = ((Val(Text4.Text) * (Sin((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If ((Val(Text4.Text) * (Cos((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = ((Val(Text4.Text) * (Cos((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If ((Val(Text4.Text) * (Cos((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = ((Val(Text4.Text) * (Cos((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= ((Val(Text4.Text) * (Cos((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N
  
Case 4
    max = ((Val(Text4.Text) * (Sin((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If ((Val(Text4.Text) * (Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = ((Val(Text4.Text) * (Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If ((Val(Text4.Text) * (Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = ((Val(Text4.Text) * (Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= ((Val(Text4.Text) * (Tan((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N
   
Case 5
    max = ((Val(Text4.Text) * (1 / Tan((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If ((Val(Text4.Text) * (1 / Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = ((Val(Text4.Text) * (1 / Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If ((Val(Text4.Text) * (1 / Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = ((Val(Text4.Text) * (1 / Tan((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= ((Val(Text4.Text) * (1 / Tan((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N
  
Case 6
    max = ((Val(Text4.Text) * (Log((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If ((Val(Text4.Text) * (Log((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = ((Val(Text4.Text) * (Log((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If ((Val(Text4.Text) * (Log((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = ((Val(Text4.Text) * (Log((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= ((Val(Text4.Text) * (Log((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N
 
Case 7
    max = ((Val(Text4.Text) * (Exp((Val(Text10.Text) * a + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    min = max
    D = (b - a) / N
    For i = a To b Step D
        If ((Val(Text4.Text) * (Exp((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) > max Then max = ((Val(Text4.Text) * (Exp((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
        If ((Val(Text4.Text) * (Exp((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) < min Then min = ((Val(Text4.Text) * (Exp((Val(Text10.Text) * i + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text)))
    Next

    DoEvents
    Randomize
    K = 0
    For i = 1 To N
    X = CSng((b - a) * Rnd() + a)
    Y = CSng((max - min) * Rnd() + min)
    If Y <= ((Val(Text4.Text) * (Exp((Val(Text10.Text) * X + Val(Text7.Text)) ^ Val(Text8.Text))) ^ Val(Text5.Text))) Then K = K + 1
    Next
    Integ = (b - a) * (max - min) * K / N

End Select

End If

str = "Интеграл равен " & Integ
Text6.Text = Text6.Text & str & vbCrLf
Frame6.Visible = True
sFile = ".\Output.txt"
    Open sFile For Output As #1
    Print #1, Text6.Text
    Close #1
End Sub

