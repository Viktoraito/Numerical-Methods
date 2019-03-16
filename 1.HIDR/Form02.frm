VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "СЛАУ методами Гаусса и итерации, нахождение обратной матрицы и определителя"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Нахождение определителя"
      Height          =   735
      Left            =   6600
      TabIndex        =   15
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   3975
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   6135
      Begin VB.Frame Frame5 
         Caption         =   "Точность"
         Height          =   975
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   1695
         Begin VB.TextBox Text4 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   13
            Text            =   "Form02.frx":0000
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Матрица"
         Height          =   2775
         Left            =   2040
         TabIndex        =   8
         Top             =   240
         Width           =   3855
         Begin VB.TextBox Text1 
            Height          =   2535
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   9
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Столбцы"
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1695
         Begin VB.TextBox Text3 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   11
            Text            =   "Form02.frx":001E
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Строки"
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1695
         Begin VB.TextBox Text2 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   10
            Text            =   "Form02.frx":003D
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Передать значения"
         Height          =   615
         Left            =   4200
         TabIndex        =   5
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Получить значения из файла InputA.txt"
         Height          =   615
         Left            =   2040
         TabIndex        =   4
         Top             =   3240
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Нахождение обратной матрицы"
      Height          =   735
      Left            =   6600
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Решение СЛАУ методом простой итерации"
      Height          =   735
      Left            =   6600
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Решение СЛАУ методом Гаусса"
      Height          =   735
      Left            =   6600
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "МАИ                                 учебная группа 3О-210Б   студент Кофман М.С."
      Height          =   615
      Index           =   1
      Left            =   8880
      TabIndex        =   14
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   9000
      Picture         =   "Form02.frx":0059
      Top             =   120
      Width           =   1770
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IsTranslated As Boolean

Private Sub Form_Load()
IsTranslated = False
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub

Private Sub Text4_Click()
Text4.Text = ""
End Sub

Private Sub Command4_Click()
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
Dim sFile As String, sWhole As String, oustr As String, outd As Double
Dim v() As String
sFile = ".\Output.txt"
    Open sFile For Output As #1
    Print #1, Text2.Text & vbCrLf & Text3.Text & vbCrLf & Text4.Text & vbCrLf & Text1.Text
    Close #1
IsTranslated = True
End Sub

Private Sub Command1_Click()
    If Text1.Text <> "" And Val(Text2.Text) <> 0 And Val(Text3.Text) <> 0 And Val(Text4.Text) <> 0 And IsTranslated Then
    Form2.Visible = True
    End If
End Sub


Private Sub Command2_Click()
    If Text1.Text <> "" And Val(Text2.Text) <> 0 And Val(Text3.Text) <> 0 And Val(Text4.Text) <> 0 And IsTranslated Then
    Form3.Visible = True
    End If
End Sub

Private Sub Command3_Click()
    If Text1.Text <> "" And Val(Text2.Text) <> 0 And Val(Text3.Text) <> 0 And Val(Text4.Text) <> 0 And IsTranslated Then
    Form4.Visible = True
    End If
End Sub


Private Sub Command6_Click()
    If Text1.Text <> "" And Val(Text2.Text) <> 0 And Val(Text3.Text) <> 0 And Val(Text4.Text) <> 0 And IsTranslated Then
    Form5.Visible = True
    End If
End Sub
