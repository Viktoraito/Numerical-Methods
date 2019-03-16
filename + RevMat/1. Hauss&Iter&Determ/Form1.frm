VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Метод Гаусса"
   ClientHeight    =   7935
   ClientLeft      =   -105
   ClientTop       =   465
   ClientWidth     =   15990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   15990
   Begin VB.Frame Frame5 
      Caption         =   "Ввод значений правой части"
      Height          =   6375
      Left            =   7800
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton Command5 
         Caption         =   "Получить значения из файла InputB.txt"
         Height          =   855
         Left            =   240
         TabIndex        =   10
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Передать значения"
         Height          =   855
         Left            =   240
         TabIndex        =   9
         Top             =   5280
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   3495
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Начальные настройки"
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2055
      Begin VB.TextBox Text6 
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Пуск"
         Height          =   735
         Left            =   240
         TabIndex        =   4
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text3 
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Строки"
         Height          =   1095
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Столбцы"
         Height          =   1095
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Frame Frame7 
         Caption         =   "Точность"
         Height          =   1095
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Окно ввода значений матрицы левой части"
      Height          =   5175
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox Text1 
         Height          =   3375
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   16
         Top             =   480
         Width           =   4575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Передать значения"
         Height          =   855
         Left            =   2760
         TabIndex        =   15
         Top             =   4080
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Получить значения из файла InputA.txt"
         Height          =   855
         Left            =   360
         TabIndex        =   14
         Top             =   4080
         Width           =   2055
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Решение"
      Height          =   3975
      Left            =   10560
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox Text2 
         Height          =   3375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   17
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Label Label1 
      Height          =   855
      Left            =   12720
      TabIndex        =   18
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   10680
      Top             =   4680
      Width           =   1770
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arr(), Barr(), X(), Check() As Double
Dim snum As Integer, cnum As Integer, eps As Integer

Private Sub Command3_Click()
snum = Val(Text3.Text)
cnum = Val(Text4.Text)
eps = Val(Text6.Text)
ReDim Barr(snum)
ReDim X(snum)
ReDim Arr(snum, cnum)
ReDim Check(snum, cnum)
Frame1.Visible = True
End Sub

Private Sub Command1_Click()
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

Text1.Text = outstr
End Sub

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

Private Sub Command2_Click()
Dim sFile As String, sWhole As String, oustr As String, outd As Double
    Dim v() As String
    Dim s As String, i As Integer, j As Integer

    sFile = ".\Output.txt"
    Open sFile For Output As #1
    Print #1, Text1.Text
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

Frame5.Visible = True
End Sub

Private Sub Command4_Click()
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

sFile = ".\Output.txt"
    Open sFile For Output As #1
    Print #1, Text2.Text
    Close #1
End Sub


Private Sub Text3_Click()
Text3.Text = ""
End Sub

Private Sub Text4_Click()
Text4.Text = ""
End Sub

Private Sub Text6_Click()
Text6.Text = ""
End Sub

