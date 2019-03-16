VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Метод прогонки"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Решение"
      Height          =   3255
      Left            =   10800
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox Text6 
         Height          =   2775
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ввод значений правой части"
      Height          =   5535
      Left            =   8040
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton Command5 
         Caption         =   "Передать значения"
         Height          =   855
         Left            =   240
         TabIndex        =   17
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Получить значения из файла InputB.txt"
         Height          =   855
         Left            =   240
         TabIndex        =   16
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   2775
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Окно ввода значений матрицы левой части"
      Height          =   4455
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton Command3 
         Caption         =   "Передать значения"
         Height          =   855
         Left            =   2400
         TabIndex        =   15
         Top             =   3360
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Получить значения из файла InputA.txt"
         Height          =   855
         Left            =   240
         TabIndex        =   14
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   2775
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2415
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   735
         Left            =   240
         TabIndex        =   13
         Top             =   4440
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "Form3.frx":0000
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Form3.frx":0022
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   975
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "Form3.frx":0043
         Top             =   600
         Width           =   1815
      End
      Begin VB.Frame Frame5 
         Caption         =   "Строки"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
      Begin VB.Frame Frame6 
         Caption         =   "Столбцы"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Frame Frame7 
         Caption         =   "Точность"
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      Caption         =   "МАИ                                 учебная группа 3О-210Б   студент Кофман М.С.      Решение СЛАУ методом прогонки"
      Height          =   1215
      Index           =   1
      Left            =   12720
      TabIndex        =   18
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   10800
      Picture         =   "Form3.frx":0062
      Top             =   3840
      Width           =   1770
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arr() As Double, Barr() As Double, X() As Double, P() As Double, Q() As Double
Dim snum As Integer, cnum As Integer, eps As Double

Private Sub Command1_Click()
snum = Val(Text1.Text)
cnum = Val(Text2.Text)
eps = Val(Text3.Text)
ReDim Barr(snum)
ReDim X(snum)
ReDim P(snum + 1)
ReDim Q(snum + 1)
ReDim Arr(snum, cnum)
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

For i = 0 To snum - 1
    For j = 0 To cnum - 1
        Arr(i, j) = 0
    Next
Next


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
If k > 0 Then
Arr(k, l + k - 1) = Val(Mid(outstr, i, j - i))
Else: Arr(k, l + k) = Val(Mid(outstr, i, j - i))
End If
l = l + 1
i = j + 1
GoTo c

B:
If k > 0 Then
Arr(k, l + k - 1) = Val(Mid(outstr, i, j - 1 - i))
Else: Arr(k, l + k) = Val(Mid(outstr, i, j - i))
End If
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

'Прогонка:
'Прямой ход с выводом
P(0) = 0
Q(0) = 0
P(1) = -Arr(0, 1) / Arr(0, 0)
Q(1) = Barr(0) / Arr(0, 0)
str = str & "P1" & "=" & Round(P(1), eps) & vbTab & vbTab & "Q1" & "=" & Round(Q(1), eps) & vbCrLf

For i = 2 To snum - 1
P(i) = -Arr(i - 1, i) / (Arr(i - 1, i - 1) + Arr(i - 1, i - 2) * P(i - 1))
Q(i) = (Barr(i - 1) - Arr(i - 1, i - 2) * Q(i - 1)) / (Arr(i - 1, i - 1) + Arr(i - 1, i - 2) * P(i - 1))
str = str & "P" & i & "=" & Round(P(i), eps) & vbTab & vbTab & "Q" & i & "=" & Round(Q(i), eps) & vbCrLf
Next
P(snum) = 0
Q(snum) = (Barr(snum - 1) - Arr(snum - 1, snum - 2) * Q(snum - 1)) / (Arr(snum - 1, snum - 1) + Arr(snum - 1, snum - 2) * P(snum - 2))
str = str & "P" & snum & "=" & Round(P(snum), eps) & vbTab & vbTab & "Q" & snum & "=" & Round(Q(snum), eps) & vbCrLf

'Обратный ход с выводом
For i = snum - 1 To 0 Step -1
X(i) = P(i + 1) * X(i + 1) + Q(i + 1)
Next

For i = 0 To snum - 1
str = str & Round(X(i), eps) & vbTab
Next

str = str & vbCrLf & vbCrLf & "Проверка (A*X):" & vbCrLf
For i = 0 To snum - 1
    k = 0
    For j = 0 To cnum - 1
        k = k + Arr(i, j) * X(j)
    Next
    str = str & k & vbCrLf
Next

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

