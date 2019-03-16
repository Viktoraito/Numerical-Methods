VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame10 
      Caption         =   "Ответ"
      Height          =   1695
      Left            =   10800
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
      Begin VB.TextBox Text8 
         Height          =   1455
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   22
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1695
      Left            =   9720
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   855
      Begin VB.CommandButton Command3 
         Caption         =   "Ответ"
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Height          =   405
         Left            =   240
         TabIndex        =   19
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "="
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ввод границ и коэффициентов"
      Height          =   1695
      Left            =   6000
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Frame Frame8 
         Caption         =   "pow"
         Height          =   615
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   615
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   0
            TabIndex        =   16
            Text            =   "1"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "k"
         Height          =   615
         Left            =   1080
         TabIndex        =   12
         Top             =   480
         Width           =   615
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   0
            TabIndex        =   13
            Text            =   "1"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ввести границы и коэффициенты в 1 измерении"
         Height          =   615
         Left            =   1800
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Caption         =   "b"
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   615
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Text            =   "1"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "a"
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   615
         Begin VB.TextBox Text3 
            Height          =   405
            Left            =   0
            TabIndex        =   8
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         Caption         =   "X1^"
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   855
         Left            =   3720
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.Frame Frame3 
         Caption         =   "Число точек"
         Height          =   975
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         Begin VB.TextBox Text2 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   5
            Text            =   "Form1.frx":0000
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Мерность"
         Height          =   975
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1455
         Begin VB.TextBox Text1 
            Height          =   735
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   4
            Text            =   "Form1.frx":0018
            Top             =   240
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dimension, N, memb As Integer, a(), b(), koeff(), p(), Eql, K, Summ, Prod As Double

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Command1_Click()
Dimension = Val(Text1.Text)
N = Val(Text2.Text)
ReDim a(N)
ReDim b(N)
ReDim koeff(N)
ReDim p(N)
ReDim X(N)
memb = 0
Frame4.Visible = True
End Sub

Private Sub Command2_Click()
If memb < Dimension - 1 Then
    a(memb) = Val(Text3.Text)
    b(memb) = Val(Text4.Text)
    koeff(memb) = Val(Text5.Text)
    p(memb) = Val(Text6.Text)
    memb = memb + 1
    Command2.Caption = "Ввести границы и коэффициенты в " & memb + 1 & " измерении"
    Label1.Caption = "X" & memb + 1 & "^"
Else
    If memb = Dimension - 1 Then
        a(memb) = Val(Text3.Text)
        b(memb) = Val(Text4.Text)
        koeff(memb) = Val(Text5.Text)
        p(memb) = Val(Text6.Text)
        Frame4.Visible = False
        Frame9.Visible = True
    End If
End If
End Sub

Private Sub Command3_Click()
Eql = Val(Text7.Text)
K = 0
DoEvents
Randomize
Prod = 1
For i = 0 To N - 1
    Summ = 0
    For j = 0 To Dimension - 1
        Summ = Summ + koeff(j) * (CSng((b(j) - a(j)) * Rnd() + a(j))) ^ p(j)
    Next
    If Summ <= Eql Then K = K + 1
Next
For i = 0 To Dimension - 1
    Prod = Prod * (b(i) - a(i))
Next
Text8.Text = Text8.Text & Prod * K / N
Frame10.Visible = True
End Sub
