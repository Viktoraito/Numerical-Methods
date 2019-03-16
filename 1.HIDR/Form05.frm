VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Определитель матрицы"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7110
   LinkTopic       =   "Form5"
   ScaleHeight     =   3960
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   " Решение"
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox Text6 
         Height          =   3015
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
      Left            =   4800
      Picture         =   "Form05.frx":0000
      Top             =   240
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "МАИ                                     учебная группа 3О-210Б  студент Кофман М.С.   Нахождение определителя матрицы"
      Height          =   975
      Left            =   4680
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim sFile, sWhole, oustr As String, outd As Double
Dim v As Variant
Dim str As String
Dim s As String, i, j, k As Integer
Dim iTemp, swap As Integer
Dim max, Temp As Double
Dim snum, cnum, eps As Integer
Dim Arr, det As Double
    
    sFile = ".\Output.txt"
    Open sFile For Input As #1
    sWhole = Input$(LOF(1), 1)
    Close #1
    v = Split(sWhole, vbNewLine & " ")

For i = 1 To 4
For Each Item In v
    If i = 1 Then snum = Val(Mid(Item, 1, 1))
    If i = 2 Then cnum = Val(Mid(Item, 4, 1))
    If i = 3 Then eps = Val(Mid(Item, 7, 1))
    If i = 4 Then outstr = outstr & Mid(Item, 10)
Next
Next

If snum <> cnum Then
    str = "Матрица не квадратная."
    GoTo Break
End If

ReDim Arr(snum, snum)
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
sFile = ".\OutputDet.txt"
    Open sFile For Output As #1
    Print #1, Text6.Text
    Close #1
End Sub
