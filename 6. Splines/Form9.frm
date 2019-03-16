VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Кубические сплайны"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   15510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "График"
      Height          =   4455
      Left            =   480
      TabIndex        =   30
      Top             =   2280
      Visible         =   0   'False
      Width           =   7215
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         FontTransparent =   0   'False
         ForeColor       =   &H00000000&
         Height          =   3975
         Left            =   120
         ScaleHeight     =   263
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   455
         TabIndex        =   31
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Решение"
      Height          =   2775
      Left            =   8280
      TabIndex        =   28
      Top             =   3960
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox Text4 
         Height          =   2175
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   29
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Ввод узлов"
      Height          =   2775
      Left            =   12480
      TabIndex        =   24
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Frame Frame9 
         Caption         =   "Узел"
         Height          =   975
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1575
         Begin VB.TextBox Text10 
            Height          =   735
            Left            =   0
            TabIndex        =   27
            Text            =   "Введите узел 1"
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ввести узел"
         Height          =   735
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ввод членов уравнения"
      Height          =   3135
      Left            =   8280
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Text            =   "0"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Text            =   "1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3120
         TabIndex        =   20
         Text            =   "1"
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option7 
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   2400
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Exp"
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   2040
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Ln"
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Ctg"
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Tg"
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cos"
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sin"
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Добавить последний член и получить ответ"
         Height          =   615
         Left            =   2040
         TabIndex        =   12
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Добавить член"
         Height          =   615
         Left            =   2040
         TabIndex        =   11
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Text            =   "1"
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "X^"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   23
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "+"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   960
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   1695
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   735
         Left            =   5640
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         Caption         =   "Точность"
         Height          =   975
         Left            =   3840
         TabIndex        =   5
         Top             =   480
         Width           =   1575
         Begin VB.TextBox Text7 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   6
            Text            =   "Form9.frx":0000
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Число узлов"
         Height          =   975
         Left            =   2040
         TabIndex        =   3
         Top             =   480
         Width           =   1575
         Begin VB.TextBox Text6 
            Height          =   615
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   4
            Text            =   "Form9.frx":0022
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Число членов"
         Height          =   975
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         Begin VB.TextBox Text1 
            Height          =   645
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   2
            Text            =   "Form9.frx":0036
            Top             =   360
            Width           =   1455
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   12480
      Picture         =   "Form9.frx":0059
      Top             =   3480
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   $"Form9.frx":0C3F
      Height          =   975
      Index           =   1
      Left            =   12240
      TabIndex        =   32
      Top             =   6120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N, NX, memb, Ni As Integer, func(), str As String, Eps, X(), F(), P(), Q(), B(), C(), D() As Double

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text10_Click()
Text10.Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub

Private Sub Text6_Click()
Text6.Text = ""
End Sub

Private Sub Text7_Click()
Text7.Text = ""
End Sub

Private Sub Command1_Click()
Frame4.Visible = False
N = Val(Text1.Text)
NX = Val(Text6.Text)
If Val(Text6.Text) < 2 Then
Frame4.Visible = True
Text4.Text = "Применение метода кубических сплайнов невозможно"
GoTo Err
End If
ReDim func(N, 5)
ReDim X(NX)
ReDim F(NX)
ReDim B(NX)
ReDim C(NX)
ReDim D(NX)
ReDim P(NX - 1)
ReDim Q(NX - 1)
memb = 0
Eps = Val(Text7.Text)
If memb = N - 1 Then
Command2.Visible = False
Command3.Visible = True
End If
Frame3.Visible = True
Err:
End Sub

Private Sub Command2_Click()

func(memb, 0) = Val(Text8.Text)
If func(memb, 0) >= 0 Then
    str = str & "+" & func(memb, 0)
    Else
    str = str & func(memb, 0)
End If
If Option1.Value = True Then
    func(memb, 1) = "S"
    str = str & "Sin("
End If
If Option2.Value = True Then
    func(memb, 1) = "C"
    str = str & "Cos("
End If
If Option3.Value = True Then
    func(memb, 1) = "T"
    str = str & "Tg("
End If
If Option4.Value = True Then
    func(memb, 1) = "N"
    str = str & "Ctg("
End If
If Option5.Value = True Then
    func(memb, 1) = "L"
    str = str & "Ln("
End If
If Option6.Value = True Then
    func(memb, 1) = "E"
    str = str & "e^("
End If
If Option7.Value = True Then
    func(memb, 1) = ""
    str = str & "("
End If
func(memb, 2) = Val(Text9.Text)
func(memb, 3) = Val(Text2.Text)
If func(memb, 3) >= 0 Then
    str = str & "(" & func(memb, 2) & "+"
Else
    str = str & "(" & func(memb, 2)
End If
str = str & func(memb, 3) & "x^"
func(memb, 4) = Text3.Text
str = str & func(memb, 4) & ")"

memb = memb + 1
If memb = N - 1 Then
Command2.Visible = False
Command3.Visible = True
End If
End Sub

Private Sub Command3_Click()

func(memb, 0) = Val(Text8.Text)
If func(memb, 0) >= 0 Then
    str = str & "+" & func(memb, 0)
    Else
    str = str & func(memb, 0)
End If
If Option1.Value = True Then
    func(memb, 1) = "S"
    str = str & "Sin("
End If
If Option2.Value = True Then
    func(memb, 1) = "C"
    str = str & "Cos("
End If
If Option3.Value = True Then
    func(memb, 1) = "T"
    str = str & "Tg("
End If
If Option4.Value = True Then
    func(memb, 1) = "N"
    str = str & "Ctg("
End If
If Option5.Value = True Then
    func(memb, 1) = "L"
    str = str & "Ln("
End If
If Option6.Value = True Then
    func(memb, 1) = "E"
    str = str & "e^("
End If
If Option7.Value = True Then
    func(memb, 1) = ""
    str = str & "("
End If
func(memb, 2) = Val(Text9.Text)
func(memb, 3) = Val(Text2.Text)
If func(memb, 3) >= 0 Then
    str = str & "(" & func(memb, 2) & "+"
Else
    str = str & "(" & func(memb, 2)
End If
str = str & func(memb, 3) & "x^"
func(memb, 4) = Text3.Text
str = str & func(memb, 4) & ")" & vbCrLf & "Узлы:" & vbCrLf

Command3.Visible = False
Frame8.Visible = True
Ni = 0
End Sub

Private Sub Command4_Click()
If Ni < NX - 1 Then
X(Ni) = Val(Text10.Text)
Ni = Ni + 1
Text10.Text = "Введите узел " & Ni + 1
Else
X(Ni) = Val(Text10.Text)
For j = 0 To NX - 1
str = str & X(j) & vbTab
Next

For i = 0 To N - 1
    F(i) = 0
Next

For i = 0 To NX - 1
    For k = 0 To N - 1
If func(k, 1) = "S" Then
    F(i) = F(i) + func(k, 0) * Sin(func(k, 2) + func(k, 3) * (X(i) ^ func(k, 4)))
End If
If func(k, 1) = "C" Then
    F(i) = F(i) + func(k, 0) * Cos(func(k, 2) + func(k, 3) * (X(i) ^ func(k, 4)))
End If
If func(k, 1) = "T" Then
    F(i) = F(i) + func(k, 0) * Tan(func(k, 2) + func(k, 3) * (X(i) ^ func(k, 4)))
End If
If func(k, 1) = "N" Then
    F(i) = F(i) + func(k, 0) * 1 / Tan(func(k, 2) + func(k, 3) * (X(i) ^ func(k, 4)))
End If
If func(k, 1) = "L" Then
    F(i) = F(i) + func(k, 0) * Log(func(k, 2) + func(k, 3) * (X(i) ^ func(k, 4)))
End If
If func(k, 1) = "E" Then
    F(i) = F(i) + func(k, 0) * Exp(func(k, 2) + func(k, 3) * (X(i) ^ func(k, 4)))
End If
If func(k, 1) = "" Then
    F(i) = F(i) + func(k, 0) * (func(k, 2) + func(k, 3) * (X(i) ^ func(k, 4)))
End If
    Next
Next

'Решение СЛАУ для нахождения коэффициентов C сплайнов методом прогонки трехдиаг. матрицы:
P(0) = 0
Q(0) = 0
C(0) = 0
If NX > 2 Then
For i = 1 To NX - 2
    P(i) = -(X(i + 1) - X(i)) / ((X(i) - X(i - 1)) * P(i - 1) + 2 * ((X(i + 1) - X(i)) + (X(i) - X(i - 1))))
    Q(i) = (6 * ((F(i + 1) - F(i)) / (X(i + 1) - X(i)) - (F(i) - F(i - 1)) / (X(i) - X(i - 1))) - (X(i) - X(i - 1)) * Q(i - 1)) / ((X(i) - X(i - 1)) * P(i - 1) + 2 * ((X(i + 1) - X(i)) + (X(i) - X(i - 1))))
Next
C(NX - 1) = (6 * ((F(NX - 1) - F(NX - 2)) / (X(NX - 1) - X(NX - 2)) - (F(NX - 2) - F(NX - 3)) / (X(NX - 2) - X(NX - 3))) - (X(NX - 1) - X(NX - 2)) * Q(NX - 2)) / (2 * ((X(NX - 1) - X(NX - 2)) + (X(NX - 2) - X(NX - 3))) + (X(NX - 2) - X(NX - 3)) * P(NX - 2))
For i = NX - 2 To 1 Step -1
    C(i) = P(i) * C(i + 1) + Q(i)
Next
End If

'Нахождение коэффициентов B и D сплайнов:
For i = NX - 1 To 1 Step -1
    B(i) = (X(i) - X(i - 1)) * (2 * C(i) + C(i - 1)) / 6 + (F(i) - F(i - 1)) / (X(i) - X(i - 1))
    D(i) = (C(i) - C(i - 1)) / (X(i) - X(i - 1))
Next

'Вывод функций сплайна:
str = str & vbCrLf & "Сплайн:" & vbCrLf
For i = 1 To NX - 1
str = str & "S(" & i + 1 & ")=" & Round(F(i), Eps) & "+" & Round(B(i), Eps) & "*(x-" & X(i) & ")+" & Round(C(i), Eps) & "*((x-" & X(i) & ")^2)/2+" & Round(D(i), Eps) & "*((x-" & X(i) & ")^3)/6" & vbCrLf
Next

'Построение графика
Picture1.ScaleMode = vbPixels
Picture1.BackColor = RGB(255, 255, 255)
  
dx = Abs((X(NX - 1) - X(0))) / Picture1.ScaleWidth
  
Dim max, min As Double
max = F(0)
min = F(NX - 1)
For k = 1 To NX - 1
      For i = X(k - 1) To X(k) - dx Step dx
          If max < F(k) + B(k) * (i - X(k)) + C(k) * ((i - X(k)) ^ 2) / 2 + D(k) * ((i - X(k)) ^ 3) / 6 Then
              max = F(k) + B(k) * (i - X(k)) + C(k) * ((i - X(k)) ^ 2) / 2 + D(k) * ((i - X(k)) ^ 3) / 6
          End If
          If min > F(k) + B(k) * (i - X(k)) + C(k) * ((i - X(k)) ^ 2) / 2 + D(k) * ((i - X(k)) ^ 3) / 6 Then
              min = F(k) + B(k) * (i - X(k)) + C(k) * ((i - X(k)) ^ 2) / 2 + D(k) * ((i - X(k)) ^ 3) / 6
          End If
      Next
Next
  
If X(0) < X(NX - 1) Then
    Picture1.Scale (X(0), max)-(X(NX - 1), min)
Else
    Picture1.Scale (X(NX - 1), max)-(X(0), min)
End If

  
Picture1.Line (X(0), 0)-(X(NX - 1), 0)
Picture1.Line (0, max)-(0, min)
For k = 1 To NX - 1
  For i = X(k - 1) To X(k) - dx Step dx
      Dim func1, func2 As Double
      func1 = F(k) + B(k) * (i - X(k)) + C(k) * ((i - X(k)) ^ 2) / 2 + D(k) * ((i - X(k)) ^ 3) / 6
      func2 = F(k) + B(k) * (i + dx - X(k)) + C(k) * ((i + dx - X(k)) ^ 2) / 2 + D(k) * ((i + dx - X(k)) ^ 3) / 6
      Picture1.Line (i, func1)-(i + dx, func2), RGB(0, 0, 0)
  Next
Next
  
Frame4.Visible = True
Text4.Text = str
Frame5.Visible = True

End If
sFile = ".\Output.txt"
    Open sFile For Output As #1
    Print #1, Text4.Text
    Close #1
End Sub

