VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Метод вращения (Якоби)"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   14505
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Ввод матрицы"
      Height          =   3615
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox Text4 
         Height          =   2055
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Получить значения из файла InputA.txt"
         Height          =   735
         Left            =   240
         TabIndex        =   4
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   735
         Left            =   2040
         TabIndex        =   3
         Top             =   2640
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Решение"
      Height          =   3615
      Left            =   7080
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox Text6 
         Height          =   2055
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   3855
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   2175
      Begin VB.Frame Frame7 
         Caption         =   "Точность"
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1935
         Begin VB.TextBox Text3 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   10
            Text            =   "Form7.frx":0000
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "Form7.frx":0020
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Frame Frame5 
         Caption         =   "Число строк/столбцов"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   11400
      Picture         =   "Form7.frx":0047
      Top             =   120
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   $"Form7.frx":0C2D
      Height          =   1575
      Index           =   1
      Left            =   11400
      TabIndex        =   12
      Top             =   2760
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arr(), tmp, ttmp As Double
Dim snum As Integer, eps As Double
Dim T(), Temp(), U(), Check() As Double
Const PI As Double = 3.14159265358979

Private Sub Command1_Click()
snum = Val(Text1.Text)
eps = Val(Text3.Text)
ReDim Arr(snum, snum)
ReDim T(snum, snum, 0)
ReDim Temp(snum, snum)
ReDim U(snum, snum)
ReDim Check(snum, snum)
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
    Dim v() As String, str As String
    Dim s As String, i As Integer, j As Integer, k As Integer, r As Integer
    Dim ii As Integer, jj As Integer
    Dim max As Double, Fi As Double
    
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

'допостроение матрицы и определение координат максимального элемента:
For i = 0 To snum - 1
    For j = 0 To snum - 1
        If j > i Then
            Arr(i, j) = Arr(j, i)
        End If
        Check(i, j) = Arr(i, j)
    Next
Next

'Метод вращения:
k = 0
tmp = 0
Do
max = Arr(0, 1)
ii = 0
jj = 1
For i = 0 To snum - 1
    For j = 0 To snum - 1
        If j > i And Abs(Arr(i, j)) > max Then
                max = Arr(i, j)
                ii = i
                jj = j
        End If
    Next
Next

'вычисление Fi:
If Abs(Arr(ii, ii) - Arr(jj, jj)) <= 10 ^ (-eps) Then
    Fi = PI / 4
Else
    Fi = 0.5 * Atn(2 * Arr(ii, jj) / (Arr(ii, ii) - Arr(jj, jj)))
End If

str = str & "Fi(" & k & ")=" & Round(Fi, eps) & vbCrLf & vbCrLf
'заполнение матрицы T:
ReDim Preserve T(snum, snum, k + 1)
For i = 0 To snum - 1
    For j = 0 To snum - 1
        If i = j Then T(i, j, k) = 1
    Next
Next
T(ii, ii, k) = Cos(Fi)
T(ii, jj, k) = -Sin(Fi)
T(jj, ii, k) = Sin(Fi)
T(jj, jj, k) = Cos(Fi)

str = str & "T(" & k & "):" & vbCrLf
For i = 0 To snum - 1
    For j = 0 To snum - 1
        str = str & Round(T(i, j, k), eps) & vbTab
    Next
    str = str & vbCrLf
Next
str = str & vbCrLf & "A*T:" & vbCrLf

'Умножение А*Т:
For i = 0 To snum - 1
    For j = 0 To snum - 1
        Temp(i, j) = 0
    Next
Next

For i = 0 To snum - 1
    For j = 0 To snum - 1
        For r = 0 To snum - 1
            Temp(i, j) = Temp(i, j) + Arr(i, r) * T(r, j, k)
        Next
        str = str & Round(Temp(i, j), eps) & vbTab
    Next
    str = str & vbCrLf
Next
            
'Умножение transp.T*(A*T):
str = str & vbCrLf & "A(" & k + 1 & ")=transp.T*A*T:" & vbCrLf

For i = 0 To snum - 1
    For j = 0 To snum - 1
        Arr(i, j) = 0
    Next
Next

For i = 0 To snum - 1
    For j = 0 To snum - 1
        For r = 0 To snum - 1
            Arr(i, j) = Arr(i, j) + T(r, i, k) * Temp(r, j)
        Next
        str = str & Round(Arr(i, j), eps) & vbTab
    Next
    str = str & vbCrLf
Next

'проверка погрешности:
k = k + 1
tmp = 0
ttmp = 0
For i = 0 To snum - 1
    For j = i + 1 To snum - 1
        tmp = tmp + Arr(i, j) ^ 2
        ttmp = ttmp + T(i, j, k - 1) ^ 2
    Next
Next
tmp = Sqr(tmp)
ttmp = Sqr(ttmp)
str = str & vbCrLf
Loop Until tmp <= (10 ^ (-eps)) Or ttmp <= 10 ^ (-eps)
'нахождение матрицы собственных векторов:
For i = 0 To snum - 1
    For j = 0 To snum - 1
        U(i, j) = T(i, j, 0)
    Next
Next

For l = 0 To k - 2
    
    For i = 0 To snum - 1
        For j = 0 To snum - 1
            Temp(i, j) = 0
        Next
    Next
    
    For i = 0 To snum - 1
        For j = 0 To snum - 1
            For r = 0 To snum - 1
                Temp(i, j) = Temp(i, j) + U(i, r) * T(r, j, l + 1)
            Next
        Next
    Next
    
    For i = 0 To snum - 1
        For j = 0 To snum - 1
            U(i, j) = Temp(i, j)
        Next
    Next
    
Next

           
'Вывод собственных значений и соотв. им собств. векторов:
For k = 0 To snum - 1
    str = str & "Lambda(" & k + 1 & ")=" & Round(Arr(k, k), eps) & vbCrLf
    For i = 0 To snum - 1
        str = str & Round(U(i, k), eps) & vbTab
    Next
    str = str & vbCrLf & vbCrLf
Next

'Проверка:
str = str & "Проверка:" & vbCrLf
For k = 0 To snum - 1
    str = str & "A*X" & vbTab & "Lambda(" & k + 1 & ")*X" & vbCrLf
    For i = 0 To snum - 1
        tmp = 0
        For j = 0 To snum - 1
            tmp = tmp + Check(i, j) * U(j, k)
        Next
        str = str & Round(tmp, eps) & vbTab & Round(Arr(k, k) * U(i, k), eps) & vbCrLf
    Next
    str = str & vbCrLf
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


