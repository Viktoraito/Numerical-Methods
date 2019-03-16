VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Определитель методом Гаусса"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   13845
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   " Решение"
      Height          =   3015
      Left            =   6960
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox Text6 
         Height          =   2295
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ввод матрицы"
      Height          =   3615
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   735
         Left            =   2040
         TabIndex        =   11
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Получить значения из файла InputA.txt"
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   2055
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Form6.frx":0000
         Top             =   720
         Width           =   1935
      End
      Begin VB.Frame Frame7 
         Caption         =   "Точность"
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   1935
         Begin VB.TextBox Text3 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   6
            Text            =   "Form6.frx":001C
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Число строк/столбцов"
         Height          =   975
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"Form6.frx":0036
      Height          =   1215
      Index           =   1
      Left            =   11280
      TabIndex        =   12
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   11400
      Picture         =   "Form6.frx":00F0
      Top             =   360
      Width           =   1770
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arr() As Double, det As Double
Dim snum As Integer, eps As Double

Private Sub Command1_Click()
snum = Val(Text1.Text)
eps = Val(Text3.Text)
ReDim Arr(snum, snum)
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
    Dim s As String, i As Integer, j As Integer, k As Integer
    Dim iTemp As Integer, swap As Integer
    Dim max As Double, Temp As Double

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

det = 1
swap = 0
For i = 0 To snum - 2
    max = Abs(Arr(i, i))
    iTemp = i
    
'Перестановка строки с максимальным элементом:
    For k = i + 1 To snum - 1
        If Abs(Arr(k, i)) > max Then
            max = Abs(Arr(k, i))
            iTemp = k
        End If
    Next
   
    If iTemp <> i Then
        For j = 0 To snum - 1
            Temp = Arr(i, j)
            Arr(i, j) = Arr(iTemp, j)
            Arr(iTemp, j) = Temp
        Next
        swap = swap + 1
    End If

'Вычисление определителя:
    For j = i + 1 To snum - 1
        If Arr(i, i) = 0 Then
            str = "Определитель методом Гаусса вычислить нельзя."
            GoTo Break
        End If
        Dim b As Double
        b = Arr(j, i) / Arr(i, i)
        For k = i To snum - 1
            Arr(j, k) = Arr(j, k) - Arr(i, k) * b
        Next
    Next
det = det * Arr(i, i)
    
    For k = 0 To snum - 1
        For j = 0 To snum - 1
            str = str & Round(Arr(k, j), eps) & vbTab
        Next
        str = str & vbCrLf
    Next
    str = str & vbCrLf

Next

det = det * Arr(snum - 1, snum - 1)
If swap Mod 2 <> 0 Then: det = det * -1
str = str & "Определитель: " & Round(det, eps)

Break:
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

