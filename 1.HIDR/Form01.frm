VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Гаусс"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8445
   LinkTopic       =   "Form2"
   ScaleHeight     =   7515
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "Решение"
      Height          =   3975
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox Text2 
         Height          =   3375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Ввод значений правой части"
      Height          =   6375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      Begin VB.TextBox Text5 
         Height          =   3495
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Передать значения"
         Height          =   855
         Left            =   240
         TabIndex        =   2
         Top             =   5280
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Получить значения из файла InputB.txt"
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   4200
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   3120
      Picture         =   "Form01.frx":0000
      Top             =   4440
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "МАИ                                 учебная группа 3О-210Б   студент Кофман М.С.      Решение СЛАУ методом Гаусса"
      Height          =   1215
      Index           =   1
      Left            =   5400
      TabIndex        =   6
      Top             =   5640
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arr(), Barr(), X(), Check() As Double
Dim snum As Integer, cnum As Integer, eps As Integer

Private Sub Command5_Click()
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

Private Sub Command4_Click()
Dim sFile, sWhole, oustr As String, outd As Double
Dim v() As String
Dim s As String, i, j As Integer
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
    If i = 3 Then eps = Val(Mid(Item, 7, 1))
    If i = 4 Then outstr = outstr & Mid(Item, 10)
Next
Next

ReDim Barr(snum)
ReDim X(snum)
ReDim Arr(snum, cnum)
ReDim Check(snum, cnum)

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

sFile = ".\OutputHauss.txt"
    Open sFile For Output As #1
    Print #1, Text5.Text
    Close #1
    
sFile = ".\OutputHauss.txt"
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

'Гаусс:
'сортировка по максимальным элементам на диагонали:
For k = 0 To snum - 1
max = Abs(Arr(k, k))
  iTemp = k
   For i = k + 1 To snum - 1
     If Abs(Arr(i, k)) > max Then
       max = Abs(Arr(i, k))
       iTemp = i
     End If
   Next
   
   For j = 0 To cnum - 1
    Temp = Arr(k, j)
    Arr(k, j) = Arr(iTemp, j)
    Arr(iTemp, j) = Temp
   Next
    Temp = Barr(k)
    Barr(k) = Barr(iTemp)
    Barr(iTemp) = Temp

'Прямой ход с выводом:
If Arr(k, k) <> 0 Then
    Temp = Arr(k, k)
    For j = 0 To cnum - 1
    Arr(k, j) = Arr(k, j) / Temp
    Next
    Barr(k) = Barr(k) / Temp
End If

    For i = k + 1 To snum - 1
        If Arr(k, k) <> 0 Then
        Temp = Arr(i, k) / Arr(k, k)
            For j = k To cnum - 1
                Arr(i, j) = Arr(i, j) - Arr(k, j) * Temp
            Next
            Barr(i) = Barr(i) - Barr(k) * Temp
        End If
    Next
    
For i = 0 To snum - 1
For j = 0 To cnum - 1
str = str & Round(Arr(i, j), eps) & vbTab
Next
str = str & vbTab & Round(Barr(i), eps) & vbCrLf
Next
str = str & vbCrLf
    
Next

'Обратный ход с выводом:
If Arr(snum - 1, snum - 1) = 0 Then
  If Barr(snum - 1) = 0 Then
  str = str & "Существует бесконечное число решений исходной системы."
  Else: str = str & "Для этой системы нет решений."
  End If
Else
    X(snum - 1) = Barr(snum - 1) / Arr(snum - 1, snum - 1)
    For i = snum - 2 To 0 Step -1
        Temp = 0
        For j = i + 1 To snum - 1
            Temp = Temp + Arr(i, j) * X(j)
        Next
        X(i) = (Barr(i) - Temp) / Arr(i, i)
    Next
    str = str & "Транспонированный вектор-ответ Х:" & vbCrLf
    For i = 0 To snum - 1
    str = str & Round(X(i), eps) & vbTab
    Next
    str = str & vbCrLf & vbCrLf & "Проверка (A*X):" & vbCrLf
    For i = 0 To snum - 1
        Barr(i) = 0
        For j = 0 To cnum - 1
            Barr(i) = Barr(i) + Check(i, j) * X(j)
        Next
        str = str & Round(Barr(i), eps) & vbCrLf
    Next
End If
Text2.Text = str

Frame6.Visible = True

sFile = ".\OutputHauss.txt"
    Open sFile For Output As #1
    Print #1, Text2.Text
    Close #1
End Sub
