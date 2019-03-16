VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Метод дихотомии"
   ClientHeight    =   5010
   ClientLeft      =   2715
   ClientTop       =   1755
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   14655
   Begin VB.Frame Frame4 
      Caption         =   "Решение"
      Height          =   2775
      Left            =   8400
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox Text4 
         Height          =   2055
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ввод членов уравнения"
      Height          =   3135
      Left            =   4560
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Text            =   "1"
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Добавить последний член и получить ответ"
         Height          =   615
         Left            =   1800
         TabIndex        =   16
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Добавить член"
         Height          =   615
         Left            =   1800
         TabIndex        =   15
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3120
         TabIndex        =   14
         Text            =   "1"
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Text            =   "1"
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option7 
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   2400
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Exp"
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   2040
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Ln"
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Ctg"
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Tg"
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cos"
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sin"
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "X^"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   13
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3855
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   2040
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "Form5.frx":0000
         Top             =   720
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Caption         =   "Точность"
         Height          =   855
         Left            =   2040
         TabIndex        =   22
         Top             =   360
         Width           =   1575
      End
      Begin VB.Frame Frame5 
         Caption         =   "Интервал корня"
         Height          =   855
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   1455
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   840
            TabIndex        =   21
            Text            =   "b"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Text            =   "a"
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   735
         Left            =   2160
         TabIndex        =   3
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Число членов"
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         Begin VB.TextBox Text1 
            Height          =   525
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   2
            Text            =   "Form5.frx":0020
            Top             =   360
            Width           =   1455
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"Form5.frx":0041
      Height          =   1215
      Index           =   1
      Left            =   12240
      TabIndex        =   25
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   12240
      Picture         =   "Form5.frx":00D2
      Top             =   600
      Width           =   1770
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N As Integer, memb As Integer, Func() As String, a As Double, b As Double, Eps As Double
Dim str As String

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub

Private Sub Text5_Click()
Text5.Text = ""
End Sub

Private Sub Text6_Click()
Text6.Text = ""
End Sub

Private Sub Text7_Click()
Text7.Text = ""
End Sub

Private Sub Command1_Click()
N = Val(Text1.Text)
ReDim Func(N, 4)
memb = 0
a = Val(Text5.Text)
b = Val(Text6.Text)
Eps = 10 ^ (-Val(Text7.Text))
If memb = N - 1 Then
Command2.Visible = False
Command3.Visible = True
End If
Frame3.Visible = True
End Sub

Private Sub Command2_Click()

Func(memb, 3) = Text8.Text
If Func(memb, 3) >= 0 Then
    str = str & "+" & Func(memb, 3)
    Else
    str = str & Func(memb, 3)
End If
If Option1.Value = True Then
    Func(memb, 0) = Func(memb, 0) & "S"
    str = str & "Sin("
End If
If Option2.Value = True Then
    Func(memb, 0) = Func(memb, 0) & "C"
    str = str & "Cos("
End If
If Option3.Value = True Then
    Func(memb, 0) = Func(memb, 0) & "T"
    str = str & "Tg("
End If
If Option4.Value = True Then
    Func(memb, 0) = Func(memb, 0) & "N"
    str = str & "Ctg("
End If
If Option5.Value = True Then
    Func(memb, 0) = Func(memb, 0) & "L"
    If a = 0 Then a = Eps
    If b = 0 Then b = Eps
    str = str & "Ln("
End If
If Option6.Value = True Then
    Func(memb, 0) = Func(memb, 0) & "E"
    str = str & "e^("
End If
If Option7.Value = True Then
    Func(memb, 0) = Func(memb, 0) & ""
    str = str & "("
End If
Func(memb, 1) = Text2.Text
str = str & Func(memb, 1) & "x^"
Func(memb, 2) = Text3.Text
str = str & Func(memb, 2) & ")"

memb = memb + 1
If memb = N - 1 Then
Command2.Visible = False
Command3.Visible = True
End If
End Sub

Private Sub Command3_Click()
Dim F As Double, Fa As Double

Func(memb, 3) = Text8.Text
If Func(memb, 3) >= 0 Then
    str = str & "+" & Func(memb, 3)
    Else
    str = str & Func(memb, 3)
End If
If Option1.Value = True Then
    Func(memb, 0) = Func(memb, 0) & "S"
    str = str & "Sin("
End If
If Option2.Value = True Then
    Func(memb, 0) = Func(memb, 0) & "C"
    str = str & "Cos("
End If
If Option3.Value = True Then
    Func(memb, 0) = Func(memb, 0) & "T"
    str = str & "Tg("
End If
If Option4.Value = True Then
    Func(memb, 0) = Func(memb, 0) & "N"
    str = str & "Ctg("
End If
If Option5.Value = True Then
    Func(memb, 0) = Func(memb, 0) & "L"
    If a = 0 Then a = Eps
    If b = 0 Then b = Eps
    str = str & "Ln("
End If
If Option6.Value = True Then
    Func(memb, 0) = Func(memb, 0) & "E"
    str = str & "e^("
End If
If Option7.Value = True Then
    Func(memb, 0) = Func(memb, 0) & ""
    str = str & "("
End If
Func(memb, 1) = Text2.Text
str = str & Func(memb, 1) & "x^"
Func(memb, 2) = Text3.Text
str = str & Func(memb, 2) & ")" & vbCrLf & vbCrLf

Do Until Abs(b - a) <= 2 * Eps
str = str & "a=" & Round(a, -Log(Eps) / Log(10)) & vbTab & "b=" & Round(b, -Log(Eps) / Log(10)) & vbCrLf
F = 0
Fa = 0
For i = 0 To memb
        If Mid(Func(i, 0), 1, 1) = "S" Then
            F = F + Func(i, 3) * Sin(Func(i, 1) * (((a + b) / 2) ^ Func(i, 2)))
            Fa = Fa + Func(i, 3) * Sin(Func(i, 1) * (a ^ Func(i, 2)))
        End If
        If Mid(Func(i, 0), 1, 1) = "C" Then
            F = F + Func(i, 3) * Cos(Func(i, 1) * (((a + b) / 2) ^ Func(i, 2)))
            Fa = Fa + Func(i, 3) * Cos(Func(i, 1) * (a ^ Func(i, 2)))
        End If
        If Mid(Func(i, 0), 1, 1) = "T" Then
            F = F + Func(i, 3) * Tan(Func(i, 1) * (((a + b) / 2) ^ Func(i, 2)))
            Fa = Fa + Func(i, 3) * Tan(Func(i, 1) * (a ^ Func(i, 2)))
        End If
        If Mid(Func(i, 0), 1, 1) = "Т" Then
            F = F + Func(i, 3) / Tan(Func(i, 1) * (((a + b) / 2) ^ Func(i, 2)))
            Fa = Fa + Func(i, 3) / Tan(Func(i, 1) * (a ^ Func(i, 2)))
        End If
        If Mid(Func(i, 0), 1, 1) = "L" Then
            F = F + Func(i, 3) * Log(Func(i, 1) * (((a + b) / 2) ^ Func(i, 2)))
            Fa = Fa + Func(i, 3) * Log(Func(i, 1) * (a ^ Func(i, 2)))
        End If
        If Mid(Func(i, 0), 1, 1) = "E" Then
            F = F + Func(i, 3) * Exp(Func(i, 1) * (((a + b) / 2) ^ Func(i, 2)))
            Fa = Fa + Func(i, 3) * Exp(Func(i, 1) * (a ^ Func(i, 2)))
        End If
        If Mid(Func(i, 0), 1, 1) = "" Then
            F = F + Func(i, 3) * Func(i, 1) * (((a + b) / 2) ^ Func(i, 2))
            Fa = Fa + Func(i, 3) * (Func(i, 1) * (a ^ Func(i, 2)))
        End If
Next

str = str & "F((a+b)/2)=" & Round(F, -Log(Eps) / Log(10)) & vbCrLf & "F(a)=" & Round(Fa, -Log(Eps) / Log(10)) & vbCrLf
If F = 0 Then GoTo Break
    If Fa * F < 0 Then
        b = (a + b) / 2
        str = str & "F(a)*F((a+b)/2)<0" & vbCrLf
    Else
        a = (a + b) / 2
        str = str & "F(a)*F((a+b)/2)>0" & vbCrLf
    End If
str = str & "b-a=" & Round(b - a, -Log(Eps) / Log(10)) & vbCrLf & vbCrLf
Loop
Break:

str = str & "x=" & Round((a + b) / 2, -Log(Eps) / Log(10)) & vbCrLf
Dim Y As Double
Y = 0
For i = 0 To memb
If Func(i, 0) = "S" Then Y = Y + Func(memb, 3) * Sin((Func(i, 1) * (a + b) / 2) ^ Func(i, 2))
If Func(i, 0) = "C" Then Y = Y + Func(memb, 3) * Cos((Func(i, 1) * (a + b) / 2) ^ Func(i, 2))
If Func(i, 0) = "T" Then Y = Y + Func(memb, 3) * Tan((Func(i, 1) * (a + b) / 2) ^ Func(i, 2))
If Func(i, 0) = "N" Then Y = Y + Func(memb, 3) / Tan((Func(i, 1) * (a + b) / 2) ^ Func(i, 2))
If Func(i, 0) = "L" Then Y = Y + Func(memb, 3) * Log((Func(i, 1) * (a + b) / 2) ^ Func(i, 2))
If Func(i, 0) = "E" Then Y = Y + Func(memb, 3) * Exp((Func(i, 1) * (a + b) / 2) ^ Func(i, 2))
If Func(i, 0) = "" Then Y = Y + Func(memb, 3) * ((Func(i, 1) * (a + b) / 2) ^ Func(i, 2))
Next
str = str & "F(x)=" & Round(Y, -Log(Eps) / Log(10) - 1)
Text4.Text = str
Frame4.Visible = True
End Sub

