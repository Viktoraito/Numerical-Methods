VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Нахождение обратной матрицы"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   14535
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Решение"
      Height          =   3735
      Left            =   7800
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox Text2 
         Height          =   2295
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Окно ввода матрицы"
      Height          =   3735
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton Command3 
         Caption         =   "Передать значения"
         Height          =   735
         Left            =   2280
         TabIndex        =   10
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Получить значения из файла InputA.txt"
         Height          =   735
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1935
      Begin VB.TextBox Text5 
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "Form4.frx":0000
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         Caption         =   "Точность"
         Height          =   975
         Left            =   240
         TabIndex        =   7
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Frame Frame5 
         Caption         =   "Столбцы"
         Height          =   975
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
         Begin VB.TextBox Text4 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   12
            Text            =   "Form4.frx":001D
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Строки"
         Height          =   975
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1455
         Begin VB.TextBox Text3 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   11
            Text            =   "Form4.frx":0036
            Top             =   240
            Width           =   1455
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   12480
      Picture         =   "Form4.frx":004C
      Top             =   600
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   $"Form4.frx":0C32
      Height          =   1095
      Left            =   12480
      TabIndex        =   14
      Top             =   3240
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arr() As Double, Arr1() As Double, Arr0() As Double, Barr() As Double
Dim snum As Integer, cnum As Integer, eps As Double

Private Sub Command1_Click()
snum = Val(Text3.Text)
cnum = Val(Text4.Text)
eps = Val(Text5.Text)
ReDim Arr(snum, cnum)
ReDim Arr1(snum, cnum)
ReDim Arr0(snum, cnum)
ReDim Barr(snum)
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

Text1.Text = outstr
End Sub


Private Sub Command3_Click()
Dim sFile As String, sWhole As String, oustr As String, outd As Double
    Dim v() As String
    Dim s As String, i As Integer, j As Integer
    Dim str As String

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
Max = Abs(Arr0(k, k))
  iTemp = k
   For i = k + 1 To snum - 1
     If Abs(Arr0(i, k)) > Max Then
       Max = Abs(Arr0(i, k))
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
Text2.Text = str

Frame3.Visible = True

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

Private Sub Text5_Click()
Text5.Text = ""
End Sub
