VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Метод простой итерации"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   14925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Решение"
      Height          =   3495
      Left            =   10200
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Text6 
         Height          =   2895
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ввод значений правой части"
      Height          =   5295
      Left            =   7440
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton Command5 
         Caption         =   "Передать значения"
         Height          =   735
         Left            =   120
         TabIndex        =   14
         Top             =   4320
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Получить значения из файла InputB.txt"
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Окно ввода значений матрицы левой части"
      Height          =   4695
      Left            =   3120
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton Command3 
         Caption         =   "Передать значения"
         Height          =   855
         Left            =   2040
         TabIndex        =   12
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Получить значения из файла InputA.txt"
         Height          =   855
         Left            =   240
         TabIndex        =   11
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   2775
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   5535
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2175
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   975
         Left            =   240
         TabIndex        =   10
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
      Begin VB.Frame Frame5 
         Caption         =   "Строки"
         Height          =   1095
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
      Begin VB.Frame Frame6 
         Caption         =   "Столбцы"
         Height          =   735
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Frame Frame7 
         Caption         =   "Точность"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   3000
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Caption         =   "МАИ                                     учебная группа 3О-210Б  студент Кофман М.С.   Решение СЛАУ методом простой итерации"
      Height          =   1095
      Left            =   12360
      TabIndex        =   18
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   10320
      Top             =   4440
      Width           =   1770
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arr(), Barr(), Arr0(), Tmp(), RevA(), X(), X1() As Double
Dim snum As Integer, cnum As Integer, eps As Double

Private Sub Command1_Click()
snum = Val(Text1.Text)
cnum = Val(Text2.Text)
eps = 1
Dim i As Integer
For i = 1 To Val(Text3.Text)
eps = eps / 10
Next
ReDim Barr(snum)
ReDim Tmp(snum)
ReDim X(snum)
ReDim X1(snum)
ReDim Arr(snum, cnum)
ReDim Arr0(snum, cnum)
ReDim RevA(snum, cnum)
Frame2.Visible = True
End Sub

Private Sub Command2_Click()
 Dim sFile As String, sWhole As String, outstr As String
    Dim v As Variant
    
    sFile = ".\InputA.txt"
    Open sFile For Input As #1
    sWhole = Input$(LOF(1), 1)
    Close #1
    v = Split(sWhole, vbNewLine & " ")

outstr = ""
For Each Item In v
outstr = outstr & Item
Next

Text4.Text = outstr
End Sub

Private Sub Command4_Click()
Dim sFile As String, sWhole As String, outstr As String
    Dim v As Variant
    
    sFile = ".\InputB.txt"
    Open sFile For Input As #1
    sWhole = Input$(LOF(1), 1)
    Close #1
    v = Split(sWhole, vbNewLine & " ")

outstr = ""
For Each Item In v
outstr = outstr & Item
Next

Text5.Text = outstr
End Sub

Private Sub Command3_Click()
Dim sFile As String, sWhole As String, oustr As String, outd As Double
    Dim v() As String
    Dim s As String, i As Integer, j As Integer

    sFile = ".\Output.txt"
    Open sFile For Output As #1
    Print #1, Text4.Text
    Close #1
    
    sFile = ".\Output.txt"
    Open sFile For Input As #1
    sWhole = Input$(LOF(1), 1)
    Close #1
    v = Split(sWhole, vbNewLine & " ")

outstr = ""
For Each Item In v
outstr = outstr & Item
Next

outstr = outstr & "\n"

Dim k As Integer
k = 0
Do While (True)
i = 1
l = 0
s = vbNullString
c:
j = i
Do While (True)
s = Mid(outstr, j, 1)
If StrComp(s, " ") = 0 Then GoTo A 'переход с j=номер крайней цифры+1
If StrComp(s, vbLf) = 0 Then GoTo B 'переход j=позиция перехода на след.строку
If StrComp(s, vbNullString) = 0 Then GoTo D
j = j + 1
Loop

A:
Arr(k, l) = Val(Mid(outstr, i, j - i))
l = l + 1
i = j + 1
GoTo c

B:
Arr(k, l) = Val(Mid(outstr, i, j - 1 - i))
l = 0
k = k + 1
i = j + 1
GoTo c
Loop
D:

Frame3.Visible = True
End Sub

Private Sub Command5_Click()

Dim sFile As String, sWhole As String, outstr As String
Dim v() As String
Dim iTemp As Integer
Dim max As Double, Temp As Double
Dim str As String

sFile = ".\Output.txt"
    Open sFile For Output As #1
    Print #1, Text5.Text
    Close #1
    
sFile = ".\Output.txt"
    Open sFile For Input As #1
    sWhole = Input$(LOF(1), 1)
    Close #1
    v = Split(sWhole, vbNewLine & " ")

outstr = ""
For Each Item In v
outstr = outstr & Item
Next

outstr = outstr & "\n"
    
Dim k As Integer
k = 0
Do While (True)
i = 1
l = 0
s = vbNullString
B:
j = i
Do While (True)
s = Mid(outstr, j, 1)
If StrComp(s, vbLf) = 0 Then GoTo A 'переход j=позиция перехода на след.строку
If StrComp(s, vbNullString) = 0 Then GoTo c
j = j + 1
Loop

A:
Barr(k) = Val(Mid(outstr, i, j - 1 - i))
k = k + 1
i = j + 1
GoTo B
Loop
c:

'Простые итерации:
'Нахождение обратной матрицы А^-1:
For k = 0 To snum - 1
    For i = 0 To snum - 1
        For j = 0 To cnum - 1
            Arr0(i, j) = Arr(i, j)
        Next
        Tmp(i) = 0
    Next
    Tmp(k) = 1

    'Cортировка по максимальным элементам на диагонали:
    max = Abs(Arr0(k, k))
    iTemp = k
    For i = k + 1 To snum - 1
        If Abs(Arr0(i, k)) > max Then
            max = Abs(Arr0(i, k))
            iTemp = i
        End If
    Next
   
    For j = 0 To cnum - 1
        Temp = Arr0(k, j)
        Arr0(k, j) = Arr0(iTemp, j)
        Arr0(iTemp, j) = Temp
    Next
    Temp = Tmp(k)
    Tmp(k) = Tmp(iTemp)
    Tmp(iTemp) = Temp

    'Прямой ход:
    If Arr0(k, k) <> 0 Then
        Temp = Arr0(k, k)
        For j = 0 To cnum - 1
            Arr0(k, j) = Arr0(k, j) / Temp
        Next
        Tmp(k) = Tmp(k) / Temp
    End If

    For i = k + 1 To snum - 1
        If Arr0(k, k) <> 0 Then
            Temp = Arr0(i, k) / Arr0(k, k)
            For j = k To cnum - 1
                Arr0(i, j) = Arr0(i, j) - Arr0(k, j) * Temp
            Next
            Tmp(i) = Tmp(i) - Tmp(k) * Temp
        End If
    Next

    'Обратный ход:
    RevA(snum - 1, k) = Tmp(snum - 1) / Arr0(snum - 1, snum - 1)
    For i = snum - 2 To 0 Step -1
        Temp = 0
        For j = i + 1 To snum - 1
            Temp = Temp + Arr0(i, j) * RevA(j, k)
        Next
        RevA(i, k) = (Tmp(i) - Temp) / Arr0(i, i)
    Next
Next
Text6.Text = str
Frame4.Visible = True
'Начальная инициализация X:
For i = 0 To snum - 1
    X(i) = Barr(i) / Arr(i, i)
Next

'Собственно, метод простых итераций с промежуточным выводом:
k = 1
While True
'X[i]=A*X[i-1]-B:
str = str & "Итерация " & k & ": "
For i = 0 To snum - 1
    X1(i) = 0
    For j = 0 To cnum - 1
        X1(i) = X1(i) + Arr(i, j) * X(j)
    Next
    X1(i) = X1(i) - Barr(i)
Next

'X[i]=X[i-1]-A^-1*(A*X[i-1]-B):
For i = 0 To snum - 1
    Tmp(i) = 0
    For j = 0 To cnum - 1
        Tmp(i) = Tmp(i) + RevA(i, j) * X1(j)
    Next
    Tmp(i) = X(i) - Tmp(i)
Next

For i = 0 To snum - 1
    X1(i) = Tmp(i)
Next

'проверка на достижение точности
norma = 0
For i = 0 To snum - 1
    If norma < Abs(X1(i) - X(i)) Then: norma = Abs(X1(i) - X(i))
Next

For i = 0 To snum - 1
X(i) = X1(i)
str = str & Round(X(i), Val(Text3.Text) + 1) & vbTab
Next
str = str & vbCrLf
k = k + 1
If norma < eps Then: GoTo Fin

Wend

Fin:
Text6.Text = str
Frame4.Visible = True
sFile = ".\Output.txt"
    Open sFile For Output As #1
    Print #1, Text6.Text
    Close #1
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

