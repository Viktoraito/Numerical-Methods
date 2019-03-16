VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Простая итерация"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7560
   LinkTopic       =   "Form3"
   ScaleHeight     =   6435
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Ввод значений правой части"
      Height          =   5295
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2415
      Begin VB.TextBox Text5 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Получить значения из файла InputB.txt"
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   3360
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Передать значения"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   4320
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Решение"
      Height          =   3495
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox Text6 
         Height          =   2895
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   3000
      Picture         =   "Form03.frx":0000
      Top             =   3960
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "МАИ                                     учебная группа 3О-210Б  студент Кофман М.С.   Решение СЛАУ методом простой итерации"
      Height          =   1095
      Left            =   5160
      TabIndex        =   6
      Top             =   4320
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arr(), Barr(), Arr0(), Tmp(), RevA(), X(), X1(), Check() As Double
Dim snum As Integer, cnum As Integer, eps As Double

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

Private Sub Command5_Click()
Dim sFile, sWhole, oustr As String, outd As Double
Dim s As String, i, j As Integer
Dim v() As String
Dim iTemp As Integer
Dim max, Temp As Double
Dim str As String

    sFile = ".\Output.txt"
    Open sFile For Input As #1
    sWhole = Input$(LOF(1), 1)
    Close #1
    v = Split(sWhole, vbNewLine & " ")

outstr = ""
For i = 1 To 4
For Each Item In v
    If i = 1 Then snum = Val(Mid(Item, 1, 1))
    If i = 2 Then cnum = Val(Mid(Item, 4, 1))
    If i = 3 Then
        eps = 1
        For j = 1 To Val(Mid(Item, 7, 1))
            eps = eps / 10
        Next
    End If
    If i = 4 Then outstr = outstr & Mid(Item, 10)
Next
Next

ReDim Barr(snum)
ReDim Tmp(snum)
ReDim X(snum)
ReDim X1(snum)
ReDim Arr(snum, cnum)
ReDim Arr0(snum, cnum)
ReDim RevA(snum, cnum)
ReDim Check(snum, cnum)

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
If StrComp(s, vbLf) = 0 Then GoTo b 'переход j=позиция перехода на след.строку
If StrComp(s, vbNullString) = 0 Then GoTo D
j = j + 1
Loop

A:
Arr(k, l) = Val(Mid(outstr, i, j - i))
l = l + 1
i = j + 1
GoTo c

b:
Arr(k, l) = Val(Mid(outstr, i, j - 1 - i))
l = 0
k = k + 1
i = j + 1
GoTo c
Loop
D:
    
sFile = ".\OutputIter.txt"
    Open sFile For Output As #1
    Print #1, Text5.Text
    Close #1
    
sFile = ".\OutputIter.txt"
    Open sFile For Input As #1
    sWhole = Input$(LOF(1), 1)
    Close #1
    v = Split(sWhole, vbNewLine & " ")

outstr = ""
For Each Item In v
outstr = outstr & Item
Next

outstr = outstr & "\n"
    
k = 0
Do While (True)
i = 1
l = 0
s = vbNullString
B2:
j = i
Do While (True)
s = Mid(outstr, j, 1)
If StrComp(s, vbLf) = 0 Then GoTo A2 'переход j=позиция перехода на след.строку
If StrComp(s, vbNullString) = 0 Then GoTo c2
j = j + 1
Loop

A2:
Barr(k) = Val(Mid(outstr, i, j - 1 - i))
k = k + 1
i = j + 1
GoTo B2
Loop
c2:

For i = 0 To snum - 1
    For j = 0 To cnum - 1
        Check(i, j) = Arr(i, j)
    Next
Next

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
str = str & Round(X(i), -Log(eps) / Log(10) + 1) & vbTab
Next
str = str & vbCrLf
k = k + 1
If norma < eps Then: GoTo Fin

Wend

Fin:
str = str & vbCrLf & "Проверка (A*X):" & vbCrLf
For i = 0 To snum - 1
    Barr(i) = 0
    For j = 0 To cnum - 1
        Barr(i) = Barr(i) + Check(i, j) * X(j)
    Next
    str = str & Round(Barr(i), -Log(eps) / Log(10)) & vbCrLf
Next
Text6.Text = str
Frame4.Visible = True
sFile = ".\OutputIter.txt"
    Open sFile For Output As #1
    Print #1, Text6.Text
    Close #1
End Sub
