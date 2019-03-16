VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Обратная матрица"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6825
   LinkTopic       =   "Form4"
   ScaleHeight     =   4320
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Решение"
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      Begin VB.TextBox Text2 
         Height          =   3135
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Label Label1 
      Caption         =   "МАИ                                     учебная группа 3О-210Б  студент Кофман М.С.   Нахождение обратной матрицы"
      Height          =   1095
      Left            =   4680
      TabIndex        =   2
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   4800
      Picture         =   "Form04.frx":0000
      Top             =   360
      Width           =   1770
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim Arr(), Arr1(), Arr0(), Barr() As Double
Dim snum, cnum As Integer, eps, Temp As Double
Dim sFile, sWhole, oustr As String, outd As Double
Dim v() As String
Dim s As String, i, j As Integer
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
    If i = 3 Then eps = Val(Mid(Item, 7, 1))
    If i = 4 Then outstr = outstr & Mid(Item, 10)
Next
Next

ReDim Arr(snum, cnum)
ReDim Arr1(snum, cnum)
ReDim Arr0(snum, cnum)
ReDim Barr(snum)

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

'Нахождение обратной матрицы:
'Задание временных матриц:
For k = 0 To snum - 1

For i = 0 To snum - 1
    For j = 0 To cnum - 1
        Arr0(i, j) = Arr(i, j)
    Next
    Barr(i) = 0
Next
Barr(k) = 1

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
    Arr0(k, j) = Arr(iTemp, j)
    Arr0(iTemp, j) = Temp
   Next
    Temp = Barr(k)
    Barr(k) = Barr(iTemp)
    Barr(iTemp) = Temp

'Прямой ход с выводом:
If Arr0(k, k) <> 0 Then
    Temp = Arr0(k, k)
    For j = 0 To cnum - 1
    Arr0(k, j) = Arr0(k, j) / Temp
    Next
    Barr(k) = Barr(k) / Temp
End If

    For i = k + 1 To snum - 1
        If Arr0(k, k) <> 0 Then
        Temp = Arr0(i, k) / Arr0(k, k)
            For j = k To cnum - 1
                Arr0(i, j) = Arr0(i, j) - Arr0(k, j) * Temp
            Next
            Barr(i) = Barr(i) - Barr(k) * Temp
        End If
    Next

'Обратный ход с выводом:
If Arr0(snum - 1, snum - 1) = 0 Then
  str = str & "Невозможно определить обратную матрицу."
Else
    Arr1(snum - 1, k) = Barr(snum - 1) / Arr0(snum - 1, snum - 1)
    For i = snum - 2 To 0 Step -1
        Temp = 0
        For j = i + 1 To snum - 1
            Temp = Temp + Arr0(i, j) * Arr1(j, k)
        Next
        Arr1(i, k) = (Barr(i) - Temp) / Arr0(i, i)
    Next
End If
Next

    For i = 0 To snum - 1
        For j = 0 To cnum - 1
            str = str & Round(Arr1(i, j), eps) & vbTab
        Next
        str = str & vbCrLf
    Next

str = str & vbCrLf & "Проверка (A^-1 * A):" & vbCrLf
For i = 0 To snum - 1
    For j = 0 To snum - 1
        Temp = 0
        For k = 0 To snum - 1
            Temp = Temp + Arr(i, k) * Arr1(k, j)
        Next
        str = str & Round(Temp, 0) & vbTab
    Next
    str = str & vbCrLf
Next

Text2.Text = str

sFile = ".\OutputRevMat.txt"
    Open sFile For Output As #1
    Print #1, Text2.Text
    Close #1

End Sub
