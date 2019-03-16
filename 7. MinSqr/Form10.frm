VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "МНК"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16035
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   16035
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Решение"
      Height          =   2655
      Left            =   11520
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox Text5 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ввод значений"
      Height          =   2655
      Left            =   7200
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox Text4 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "Form10.frx":0000
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   11
         Text            =   "Form10.frx":0022
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Добавить значения и получить ответ"
         Height          =   855
         Left            =   2160
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Добавить значения"
         Height          =   855
         Left            =   2160
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   1455
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   735
         Left            =   3240
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Точность"
         Height          =   855
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         Begin VB.TextBox Text2 
            Height          =   615
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   4
            Text            =   "Form10.frx":0040
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Число Узлов"
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
         Begin VB.TextBox Text1 
            Height          =   615
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   3
            Text            =   "Form10.frx":0060
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"Form10.frx":0074
      Height          =   975
      Index           =   1
      Left            =   2280
      TabIndex        =   13
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   240
      Picture         =   "Form10.frx":0111
      Top             =   240
      Width           =   1770
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Eps, X(), Y(), A1(2, 2), A2(3, 3), B1(2), B2(3), Alpha1(2), Alpha2(3), Summ As Double, NX, memb As Integer, str As String

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

Private Sub Command1_Click()
Frame5.Visible = False
NX = Val(Text1.Text)
ReDim X(NX)
ReDim Y(NX)
If Val(Text2.Text) < 2 Then
Frame5.Visible = True
Text5.Text = "Применение метода наименьших квадратов невозможно"
GoTo Err
End If
Eps = Val(Text2.Text)
Frame4.Visible = True
memb = 0
Err:
End Sub

Private Sub Command2_Click()
X(memb) = Val(Text3.Text)
Y(memb) = Val(Text4.Text)
memb = memb + 1
Text3.Text = "Введите значение в узле " & memb + 1
Text4.Text = "Введите значение функции в узле " & memb + 1
If memb = NX - 1 Then
Command2.Visible = False
Command3.Visible = True
End If
End Sub

Private Sub Command3_Click()
X(memb) = Val(Text3.Text)
Y(memb) = Val(Text4.Text)
Command3.Visible = False

For i = 0 To 1
    For j = 0 To 1
        A1(i, j) = 0
    Next
    B1(i) = 0
Next

For i = 0 To 2
    For j = 0 To 2
        A2(i, j) = 0
    Next
    B2(i) = 0
Next

A1(0, 0) = NX
A2(0, 0) = NX
For i = 0 To NX - 1
    A2(0, 1) = A2(0, 1) + X(i)
    A2(1, 1) = A2(1, 1) + X(i) ^ 2
    A2(1, 2) = A2(1, 2) + X(i) ^ 3
    A2(2, 2) = A2(2, 2) + X(i) ^ 4
    B2(0) = B2(0) + Y(i)
    B2(1) = B2(1) + Y(i) * X(i)
    B2(2) = B2(2) + Y(i) * X(i) ^ 2
Next
A1(0, 1) = A2(0, 1)
A1(1, 0) = A2(0, 1)
A1(1, 1) = A2(1, 1)
A2(0, 2) = A2(1, 1)
A2(1, 0) = A2(0, 1)
A2(2, 0) = A2(1, 1)
A2(2, 1) = A2(1, 2)
B1(0) = B2(0)
B1(1) = B2(1)

'Нахождение линейного многочлена:
Temp = A1(0, 0)
For j = 0 To 1
    A1(0, j) = A1(0, j) / Temp
Next
B1(0) = B1(0) / Temp
Temp = A1(1, 0) / A1(0, 0)
For j = 0 To 1
    A1(1, j) = A1(1, j) - A1(0, j) * Temp
Next
B1(1) = B1(1) - B1(0) * Temp
    
str = "Линейный многочлен:" & vbCrLf
Alpha1(1) = B1(1) / A1(1, 1)
Alpha1(0) = (B1(0) - A1(0, 1) * Alpha1(1)) / A1(0, 0)
str = str & Round(Alpha1(0), Eps) & "+" & Round(Alpha1(1), Eps) & "*x" & vbCrLf

Summ = 0
For i = 0 To NX - 1
        Summ = Summ + (Y(i) - (Alpha1(0) + Alpha1(1) * X(i))) ^ 2
Next
str = str & "Сумма квадратов отклонений в узлах: " & Round(Summ, Eps) & vbCrLf

'Нахождение квадратичного многочлена:
For k = 0 To 2
Max = Abs(A2(k, k))
  iTemp = k
   For i = k + 1 To 2
     If Abs(A2(i, k)) > Max Then
       Max = Abs(A2(i, k))
       iTemp = i
     End If
   Next
   
   For j = 0 To 2
    Temp = A2(k, j)
    A2(k, j) = A2(iTemp, j)
    A2(iTemp, j) = Temp
   Next
    Temp = B2(k)
    B2(k) = B2(iTemp)
    B2(iTemp) = Temp

If A2(k, k) <> 0 Then
    Temp = A2(k, k)
    For j = 0 To 2
    A2(k, j) = A2(k, j) / Temp
    Next
    B2(k) = B2(k) / Temp
End If

    For i = k + 1 To 2
        If A2(k, k) <> 0 Then
        Temp = A2(i, k) / A2(k, k)
            For j = k To 2
                A2(i, j) = A2(i, j) - A2(k, j) * Temp
            Next
            B2(i) = B2(i) - B2(k) * Temp
        End If
    Next
Next
   
str = str & vbCrLf & "Квадратичный многочлен:" & vbCrLf
Alpha2(2) = B2(2) / A2(2, 2)
    For i = 1 To 0 Step -1
        Temp = 0
        For j = i + 1 To 2
            Temp = Temp + A2(i, j) * Alpha2(j)
        Next
        Alpha2(i) = (B2(i) - Temp) / A2(i, i)
    Next
str = str & Round(Alpha2(0), Eps) & "+" & Round(Alpha2(1), Eps) & "*x+" & Round(Alpha2(2), Eps) & "*x^2" & vbCrLf

Summ = 0
For i = 0 To NX - 1
        Summ = Summ + (Y(i) - (Alpha2(0) + Alpha2(1) * X(i) + Alpha2(2) * X(i) ^ 2)) ^ 2
Next
str = str & "Сумма квадратов отклонений в узлах: " & Round(Summ, Eps) & vbCrLf

Text5.Text = str
Frame5.Visible = True
sFile = ".\Output.txt"
    Open sFile For Output As #1
    Print #1, Text5.Text
    Close #1
End Sub
