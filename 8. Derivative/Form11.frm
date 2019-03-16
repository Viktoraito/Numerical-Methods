VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Численное дифференцирование"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      Caption         =   "Решение"
      Height          =   1935
      Left            =   2880
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox Text6 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   15
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Ввод узлов"
      Height          =   3495
      Left            =   600
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
      Begin VB.TextBox Text7 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   16
         Text            =   "Form11.frx":0000
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Передать последние значения и получить ответ"
         Height          =   855
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Передать значения"
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "Form11.frx":0022
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   2895
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   855
         Left            =   3840
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.Frame Frame5 
         Caption         =   "Точка"
         Height          =   975
         Left            =   2160
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
         Begin VB.TextBox Text4 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   9
            Text            =   "Form11.frx":0033
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Шаг"
         Height          =   975
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
         Begin VB.TextBox Text3 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   8
            Text            =   "Form11.frx":004F
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Точность"
         Height          =   975
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         Begin VB.TextBox Text2 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   7
            Text            =   "Form11.frx":0068
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Число узлов"
         Height          =   975
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1575
         Begin VB.TextBox Text1 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   6
            Text            =   "Form11.frx":008A
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"Form11.frx":009E
      Height          =   1095
      Index           =   1
      Left            =   2880
      TabIndex        =   17
      Top             =   5760
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   5880
      Picture         =   "Form11.frx":01FE
      Top             =   600
      Width           =   1770
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X(), Y(), F(), Eps, XX, H, DY, DDY As Double, N, memb As Integer, str As String
Dim YLeft1(), YLeft2(), YRight1(), YRight2(), YMid(), YL1, YL2, YR1, YR2, YM As Double

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

Private Sub Command1_Click()
N = Val(Text1.Text)
Eps = Val(Text2.Text)
H = Val(Text3.Text)
XX = Val(Text4.Text)
ReDim X(N)
ReDim Y(N)
ReDim F(N, N)
ReDim YLeft1(N)
ReDim YRight1(N)
ReDim YLeft2(N)
ReDim YRight2(N)
ReDim YMid(N)
If N = 1 Then
Command2.Visible = False
Command3.Visible = True
End If
If N < 1 Then
Frame7.Visible = True
Text6.Text = "Неверное число узлов"
GoTo Fin
End If
Frame6.Visible = True
memb = 0
Fin:
End Sub

Private Sub Command2_Click()
X(memb) = Val(Text5.Text)
Y(memb) = Val(Text7.Text)
memb = memb + 1
If memb = N - 1 Then
Command2.Visible = False
Command3.Visible = True
End If
Text5.Text = "Введите узел " & memb + 1
Text7.Text = "Введите значение функции в узле " & memb + 1
End Sub

Private Sub Command3_Click()
str = ""
X(memb) = Val(Text5.Text)
Y(memb) = Val(Text7.Text)
DY = Y(0)

'Вычисление коэффициентов интерп. многочлена Ньютона:
For i = 0 To N - 1
    For j = 0 To N - 1
        F(i, j) = 0
    Next
Next

For i = 0 To N - 1
F(i, 0) = Y(i)
Next

For j = 1 To N - 1
    For i = j To N - 1
        F(i, j) = F(i, j) + (F(i, j - 1) - F(i - 1, j - 1)) / (X(i) - X(i - j))
    Next
Next

'вычисление множителей интерп. многочлена Ньютона для точек x, x(+/-)h, x(+/-)2h:
YL1 = F(0, 0)
YR1 = F(0, 0)
YL2 = F(0, 0)
YR2 = F(0, 0)
YM = F(0, 0)

YLeft1(0) = XX - H - X(0)
YRight1(0) = XX + H - X(0)
YLeft2(0) = XX - 2 * H - X(0)
YRight2(0) = XX + 2 * H - X(0)
YMid(0) = XX - X(0)

For i = 1 To N - 1
    YLeft1(i) = YLeft1(i - 1) * (XX - H - X(i))
    YRight1(i) = YRight1(i - 1) * (XX + H - X(i))
    YLeft2(i) = YLeft2(i - 1) * (XX - 2 * H - X(i))
    YRight2(i) = YRight2(i - 1) * (XX + 2 * H - X(i))
    YMid(i) = YMid(i - 1) * (XX - X(i))
Next

For i = 1 To N - 1
    YL1 = YL1 + F(i, i) * YLeft1(i - 1)
    YR1 = YR1 + F(i, i) * YRight1(i - 1)
    YL2 = YL2 + F(i, i) * YLeft2(i - 1)
    YR2 = YR2 + F(i, i) * YRight2(i - 1)
    YM = YM + F(i, i) * YMid(i - 1)
Next

str = str & "y'(" & XX & ")=(" & Round(YR1, Eps) & "-" & Round(YL1, Eps) & ")/(2*" & H & ")" & vbCrLf
str = str & "y''(" & XX & ")=(" & Round(YR2, Eps) & "-2*" & Round(YM, Eps) & "+" & Round(YL2, Eps) & ")/(4*" & H & "^2)" & vbCrLf
str = str & "Первая производная в точке " & Round(XX, Eps) & " равна " & Round((YR1 - YL1) / (2 * H), Eps) & vbCrLf & "Вторая производная равна " & Round((YR2 - 2 * YM + YL2) / ((2 * H) ^ 2), Eps)
Text6.Text = str
Frame7.Visible = True
sFile = ".\Output.txt"
    Open sFile For Output As #1
    Print #1, Text6.Text
    Close #1
End Sub
