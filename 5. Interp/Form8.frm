VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Интерполяционные многочлены "
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   15090
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Caption         =   "Ввод узлов"
      Height          =   2775
      Left            =   8040
      TabIndex        =   29
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CommandButton Command4 
         Caption         =   "Ввести узел"
         Height          =   735
         Left            =   240
         TabIndex        =   32
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Frame Frame9 
         Caption         =   "Узел"
         Height          =   975
         Left            =   240
         TabIndex        =   30
         Top             =   480
         Width           =   1695
         Begin VB.TextBox Text10 
            Height          =   735
            Left            =   0
            TabIndex        =   31
            Text            =   "Введите узел 1"
            Top             =   240
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   3495
      Left            =   120
      TabIndex        =   16
      Top             =   360
      Width           =   3855
      Begin VB.Frame Frame7 
         Caption         =   "Число узлов"
         Height          =   855
         Left            =   2040
         TabIndex        =   26
         Top             =   360
         Width           =   1575
         Begin VB.TextBox Text6 
            Height          =   615
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   27
            Text            =   "Form8.frx":0000
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Число членов"
         Height          =   855
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1455
         Begin VB.TextBox Text1 
            Height          =   525
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   22
            Text            =   "Form8.frx":0014
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   735
         Left            =   240
         TabIndex        =   20
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         Caption         =   "Точка проверки"
         Height          =   975
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
         Begin VB.TextBox Text5 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   19
            Text            =   "Form8.frx":0037
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Точность"
         Height          =   975
         Left            =   2040
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
         Begin VB.TextBox Text7 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   28
            Text            =   "Form8.frx":005F
            Top             =   240
            Width           =   1575
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ввод членов уравнения"
      Height          =   3135
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Text            =   "0"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         TabIndex        =   7
         Text            =   "1"
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Добавить член"
         Height          =   615
         Left            =   2040
         TabIndex        =   5
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Добавить последний член и получить ответ"
         Height          =   615
         Left            =   2040
         TabIndex        =   4
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sin"
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cos"
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Tg"
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Ctg"
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Ln"
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Exp"
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   2040
         Width           =   855
      End
      Begin VB.OptionButton Option7 
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   2400
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3120
         TabIndex        =   6
         Text            =   "1"
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "+"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "X^"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   15
         Top             =   960
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Решение"
      Height          =   2775
      Left            =   10440
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox Text4 
         Height          =   2055
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   10440
      Picture         =   "Form8.frx":0081
      Top             =   3240
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   $"Form8.frx":0C67
      Height          =   1455
      Index           =   1
      Left            =   12360
      TabIndex        =   23
      Top             =   3480
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N, NX, memb, Ni As Integer, Func(), str As String, XX, Eps, Err, Err1, Ferr, Err2(), X(), F() As Double, Delt() As String
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
NX = Val(Text6.Text)
ReDim Func(N, 5)
ReDim X(NX)
ReDim F(NX, NX)
ReDim Delt(NX)
ReDim Err2(NX)
memb = 0
XX = Val(Text5.Text)
Eps = Val(Text7.Text)
If memb = N - 1 Then
Command2.Visible = False
Command3.Visible = True
End If
Frame3.Visible = True
End Sub

Private Sub Command2_Click()

Func(memb, 0) = Val(Text8.Text)
If Func(memb, 0) >= 0 Then
    str = str & "+" & Func(memb, 0)
    Else
    str = str & Func(memb, 0)
End If
If Option1.Value = True Then
    Func(memb, 1) = "S"
    str = str & "Sin("
End If
If Option2.Value = True Then
    Func(memb, 1) = "C"
    str = str & "Cos("
End If
If Option3.Value = True Then
    Func(memb, 1) = "T"
    str = str & "Tg("
End If
If Option4.Value = True Then
    Func(memb, 1) = "N"
    str = str & "Ctg("
End If
If Option5.Value = True Then
    Func(memb, 1) = "L"
    str = str & "Ln("
End If
If Option6.Value = True Then
    Func(memb, 1) = "E"
    str = str & "e^("
End If
If Option7.Value = True Then
    Func(memb, 1) = ""
    str = str & "("
End If
Func(memb, 2) = Val(Text9.Text)
Func(memb, 3) = Val(Text2.Text)
If Func(memb, 3) >= 0 Then
    str = str & "(" & Func(memb, 2) & "+"
Else
    str = str & "(" & Func(memb, 2)
End If
str = str & Func(memb, 3) & "x^"
Func(memb, 4) = Text3.Text
str = str & Func(memb, 4) & ")"

memb = memb + 1
If memb = N - 1 Then
Command2.Visible = False
Command3.Visible = True
End If
End Sub

Private Sub Command3_Click()

Func(memb, 0) = Val(Text8.Text)
If Func(memb, 0) >= 0 Then
    str = str & "+" & Func(memb, 0)
    Else
    str = str & Func(memb, 0)
End If
If Option1.Value = True Then
    Func(memb, 1) = "S"
    str = str & "Sin("
End If
If Option2.Value = True Then
    Func(memb, 1) = "C"
    str = str & "Cos("
End If
If Option3.Value = True Then
    Func(memb, 1) = "T"
    str = str & "Tg("
End If
If Option4.Value = True Then
    Func(memb, 1) = "N"
    str = str & "Ctg("
End If
If Option5.Value = True Then
    Func(memb, 1) = "L"
    str = str & "Ln("
End If
If Option6.Value = True Then
    Func(memb, 1) = "E"
    str = str & "e^("
End If
If Option7.Value = True Then
    Func(memb, 1) = ""
    str = str & "("
End If
Func(memb, 2) = Val(Text9.Text)
Func(memb, 3) = Val(Text2.Text)
If Func(memb, 3) >= 0 Then
    str = str & "(" & Func(memb, 2) & "+"
Else
    str = str & "(" & Func(memb, 2)
End If
str = str & Func(memb, 3) & "x^"
Func(memb, 4) = Text3.Text
str = str & Func(memb, 4) & ")" & vbCrLf & "Узлы:" & vbCrLf

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

Dim LnX As String, Prod As Double

Ferr = 0
For k = 0 To N - 1
If Func(k, 1) = "S" Then
    Ferr = Ferr + Func(k, 0) * Sin(Func(k, 2) + Func(k, 3) * (XX ^ Func(k, 4)))
End If
If Func(k, 1) = "C" Then
    Ferr = Ferr + Func(k, 0) * Cos(Func(k, 2) + Func(k, 3) * (XX ^ Func(k, 4)))
End If
If Func(k, 1) = "T" Then
    Ferr = Ferr + Func(k, 0) * Tan(Func(k, 2) + Func(k, 3) * (XX ^ Func(k, 4)))
End If
If Func(k, 1) = "N" Then
    Ferr = Ferr + Func(k, 0) * 1 / Tan(Func(k, 2) + Func(k, 3) * (XX ^ Func(k, 4)))
End If
If Func(k, 1) = "L" Then
    Ferr = Ferr + Func(k, 0) * Log(Func(k, 2) + Func(k, 3) * (XX ^ Func(k, 4)))
End If
If Func(k, 1) = "E" Then
    Ferr = Ferr + Func(k, 0) * Exp(Func(k, 2) + Func(k, 3) * (XX ^ Func(k, 4)))
End If
If Func(k, 1) = "" Then
    Ferr = Ferr + Func(k, 0) * (Func(k, 2) + Func(k, 3) * (XX ^ Func(k, 4)))
End If
Next

LnX = ""
Err = 0
For i = 0 To NX - 1
    LnX = LnX & "+"
    Prod = 1
    For k = 0 To N - 1
If Func(k, 1) = "S" Then
    Prod = Prod * Func(k, 0) * Sin(Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
End If
If Func(k, 1) = "C" Then
    Prod = Prod * Func(k, 0) * Cos(Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
End If
If Func(k, 1) = "T" Then
    Prod = Prod * Func(k, 0) * Tan(Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
End If
If Func(k, 1) = "N" Then
    Prod = Prod * Func(k, 0) * 1 / Tan(Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
End If
If Func(k, 1) = "L" Then
    Prod = Prod * Func(k, 0) * Log(Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
End If
If Func(k, 1) = "E" Then
    Prod = Prod * Func(k, 0) * Exp(Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
End If
If Func(k, 1) = "" Then
    Prod = Prod * Func(k, 0) * (Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
End If
Next

LnX = LnX & Round(Prod, Eps)
Err1 = Prod
    
    For j = 0 To NX - 1
        If i <> j Then
            LnX = LnX & "(x-" & X(j) & ")"
            Err1 = Err1 * (XX - X(j))
        End If
    Next
    
    LnX = LnX & "/"
    
    Prod = 1
    For j = 0 To NX - 1
        If i <> j Then: Prod = Prod * (X(i) - X(j))
    Next
    
    LnX = LnX & Round(Prod, Eps)
    Err = Err + Err1 / Prod
Next
str = str & vbCrLf & "Многочлен в форме Лагранжа:" & vbCrLf & LnX & vbCrLf & "Погрешность в точке " & XX & ": " & Round(Abs(Err - Ferr), Eps) & vbCrLf

Dim PnX As String
PnX = ""
Err = 0
For i = 0 To NX - 1
    For j = 0 To NX - 1
        F(i, j) = 0
    Next
Next

For k = 0 To N - 1
If Func(k, 1) = "S" Then
    For i = 0 To NX - 1
    F(i, 0) = F(i, 0) + Func(k, 0) * Sin(Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
    Next
End If
If Func(k, 1) = "C" Then
    For i = 0 To NX - 1
    F(i, 0) = F(i, 0) + Func(k, 0) * Cos(Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
    Next
End If
If Func(k, 1) = "T" Then
    For i = 0 To NX - 1
    F(i, 0) = F(i, 0) + Func(k, 0) * Tan(Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
    Next
End If
If Func(k, 1) = "N" Then
    For i = 0 To NX - 1
    F(i, 0) = F(i, 0) + Func(k, 0) * 1 / Tan(Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
    Next
End If
If Func(k, 1) = "L" Then
    For i = 0 To NX - 1
    F(i, 0) = F(i, 0) + Func(k, 0) * Log(Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
    Next
End If
If Func(k, 1) = "E" Then
    For i = 0 To NX - 1
    F(i, 0) = F(i, 0) + Func(k, 0) * Exp(Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
    Next
End If
If Func(k, 1) = "" Then
    For i = 0 To NX - 1
    F(i, 0) = F(i, 0) + Func(k, 0) * (Func(k, 2) + Func(k, 3) * (X(i) ^ Func(k, 4)))
    Next
End If
Next

PnX = PnX & Round(F(0, 0), Eps)
Err = F(0, 0)

For k = 0 To N - 1
    For j = 1 To NX - 1
        For i = j To NX - 1
            F(i, j) = F(i, j) + (F(i, j - 1) - F(i - 1, j - 1)) / (X(i) - X(i - j))
        Next
    Next
Next

Delt(0) = "(x-" & X(0) & ")"
Err2(0) = XX - X(0)
For i = 1 To NX - 1
    Delt(i) = Delt(i - 1) & "(x-" & X(i) & ")"
    Err2(i) = Err2(i - 1) * (XX - X(i))
Next


For i = 1 To NX - 1
    PnX = PnX & "+(" & Round(F(i, i), Eps) & ")" & Delt(i - 1)
    Err = Err + F(i, i) * Err2(i - 1)
Next

str = str & "Многочлен в форме Ньютона:" & vbCrLf & PnX & vbCrLf & "Погрешность в точке " & XX & ": " & Round(Abs(Err - Ferr), Eps) & vbCrLf
Text4.Text = str
Frame4.Visible = True
sFile = ".\Output.txt"
    Open sFile For Output As #1
    Print #1, Text4.Text
    Close #1
End If
End Sub
