VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Метод Конечных Разностей"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   15795
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Caption         =   "График"
      Height          =   3015
      Left            =   12120
      TabIndex        =   59
      Top             =   240
      Width           =   3615
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   2655
         Left            =   120
         ScaleHeight     =   173
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   221
         TabIndex        =   60
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Решение"
      Height          =   3015
      Left            =   10080
      TabIndex        =   57
      Top             =   240
      Width           =   1815
      Begin VB.TextBox Text15 
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   58
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9615
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   915
         Left            =   6960
         TabIndex        =   6
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Frame Frame2 
         Caption         =   "Шаг и точность"
         Height          =   1095
         Left            =   4440
         TabIndex        =   1
         Top             =   3840
         Width           =   2295
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1680
            TabIndex        =   3
            Text            =   "1"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   360
            TabIndex        =   2
            Text            =   "0.1"
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "знаков после запятой"
            Height          =   495
            Left            =   960
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "h="
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Система"
         Height          =   5175
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   9255
         Begin VB.Frame Frame4 
            Height          =   2775
            Left            =   1080
            TabIndex        =   41
            Top             =   360
            Width           =   855
            Begin VB.OptionButton Option1 
               Caption         =   "1"
               Height          =   195
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Value           =   -1  'True
               Width           =   375
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Sin"
               Height          =   195
               Left            =   120
               TabIndex        =   47
               Top             =   600
               Width           =   615
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Cos"
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   960
               Width           =   615
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Tg"
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   1320
               Width           =   615
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Ctg"
               Height          =   255
               Left            =   120
               TabIndex        =   44
               Top             =   1680
               Width           =   615
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Ln"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   2040
               Width           =   615
            End
            Begin VB.OptionButton Option7 
               Caption         =   "Exp"
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   2400
               Width           =   615
            End
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   480
            TabIndex        =   40
            Text            =   "1"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   3600
            TabIndex        =   39
            Text            =   "1"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1920
            TabIndex        =   38
            Text            =   "1"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   2760
            TabIndex        =   37
            Text            =   "1"
            Top             =   1320
            Width           =   495
         End
         Begin VB.Frame Frame5 
            Height          =   2775
            Left            =   4200
            TabIndex        =   29
            Top             =   360
            Width           =   855
            Begin VB.OptionButton Option8 
               Caption         =   "1"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   240
               Value           =   -1  'True
               Width           =   495
            End
            Begin VB.OptionButton Option9 
               Caption         =   "Sin"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   600
               Width           =   615
            End
            Begin VB.OptionButton Option10 
               Caption         =   "Cos"
               Height          =   255
               Left            =   120
               TabIndex        =   34
               Top             =   960
               Width           =   615
            End
            Begin VB.OptionButton Option11 
               Caption         =   "Tg"
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   1320
               Width           =   615
            End
            Begin VB.OptionButton Option12 
               Caption         =   "Ctg"
               Height          =   195
               Left            =   120
               TabIndex        =   32
               Top             =   1680
               Width           =   615
            End
            Begin VB.OptionButton Option13 
               Caption         =   "Ln"
               Height          =   195
               Left            =   120
               TabIndex        =   31
               Top             =   2040
               Width           =   495
            End
            Begin VB.OptionButton Option14 
               Caption         =   "Exp"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   2400
               Width           =   615
            End
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   5040
            TabIndex        =   28
            Text            =   "1"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   5880
            TabIndex        =   27
            Text            =   "1"
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   6600
            TabIndex        =   26
            Text            =   "1"
            Top             =   1680
            Width           =   495
         End
         Begin VB.Frame Frame6 
            Height          =   2775
            Left            =   7080
            TabIndex        =   18
            Top             =   360
            Width           =   855
            Begin VB.OptionButton Option15 
               Caption         =   "1"
               Height          =   195
               Left            =   120
               TabIndex        =   25
               Top             =   240
               Value           =   -1  'True
               Width           =   495
            End
            Begin VB.OptionButton Option16 
               Caption         =   "Sin"
               Height          =   195
               Left            =   120
               TabIndex        =   24
               Top             =   600
               Width           =   615
            End
            Begin VB.OptionButton Option17 
               Caption         =   "Cos"
               Height          =   195
               Left            =   120
               TabIndex        =   23
               Top             =   960
               Width           =   615
            End
            Begin VB.OptionButton Option18 
               Caption         =   "Tg"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   1320
               Width           =   615
            End
            Begin VB.OptionButton Option19 
               Caption         =   "Ctg"
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   1680
               Width           =   615
            End
            Begin VB.OptionButton Option20 
               Caption         =   "Ln"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   2040
               Width           =   495
            End
            Begin VB.OptionButton Option21 
               Caption         =   "Exp"
               Height          =   195
               Left            =   120
               TabIndex        =   19
               Top             =   2400
               Width           =   615
            End
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   7920
            TabIndex        =   17
            Text            =   "1"
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   8640
            TabIndex        =   16
            Text            =   "1"
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   240
            TabIndex        =   15
            Text            =   "1"
            Top             =   3480
            Width           =   495
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   960
            TabIndex        =   14
            Text            =   "0"
            Top             =   3480
            Width           =   495
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   1680
            TabIndex        =   13
            Text            =   "1"
            Top             =   3480
            Width           =   495
         End
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   3120
            TabIndex        =   12
            Text            =   "1"
            Top             =   3480
            Width           =   495
         End
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   240
            TabIndex        =   11
            Text            =   "1"
            Top             =   4320
            Width           =   495
         End
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   960
            TabIndex        =   10
            Text            =   "1"
            Top             =   4320
            Width           =   495
         End
         Begin VB.TextBox Text19 
            Height          =   285
            Left            =   1680
            TabIndex        =   9
            Text            =   "1"
            Top             =   4320
            Width           =   495
         End
         Begin VB.TextBox Text21 
            Height          =   285
            Left            =   3120
            TabIndex        =   8
            Text            =   "0"
            Top             =   4320
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "1"
            Height          =   255
            Left            =   2400
            TabIndex        =   56
            Top             =   4320
            Width           =   495
         End
         Begin VB.Label Label9 
            Caption         =   "0"
            Height          =   255
            Left            =   2400
            TabIndex        =   55
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   $"Form14.frx":0000
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   1680
            Width           =   8415
         End
         Begin VB.Label Label4 
            Caption         =   "^"
            Height          =   135
            Left            =   2640
            TabIndex        =   53
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label Label5 
            Caption         =   "^"
            Height          =   135
            Left            =   5760
            TabIndex        =   52
            Top             =   1560
            Width           =   135
         End
         Begin VB.Label Label6 
            Caption         =   "^"
            Height          =   135
            Left            =   8520
            TabIndex        =   51
            Top             =   1560
            Width           =   135
         End
         Begin VB.Label Label7 
            Caption         =   " Y'(           ) +             Y(            )="
            Height          =   255
            Left            =   720
            TabIndex        =   50
            Top             =   3480
            Width           =   2415
         End
         Begin VB.Label Label8 
            Caption         =   " Y'(           ) +             Y(            )="
            Height          =   255
            Left            =   720
            TabIndex        =   49
            Top             =   4320
            Width           =   2415
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"Form14.frx":00B5
      Height          =   1455
      Index           =   1
      Left            =   12120
      TabIndex        =   61
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2400
      Index           =   1
      Left            =   10080
      Picture         =   "Form14.frx":01BB
      Top             =   3360
      Width           =   1770
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arr(), Barr(), X(), P(), Q(), a, b, h, Eps, R(), E(), Z(), Y() As Double, N As Integer
Dim Alpha, Beta, Hamma, Delta, Fi, Psi As Double, str As String

Private Sub Text1_CLick()
Text1.Text = ""
End Sub
Private Sub Text2_CLick()
Text2.Text = ""
End Sub
Private Sub Text3_CLick()
Text3.Text = ""
End Sub
Private Sub Text4_CLick()
Text4.Text = ""
End Sub
Private Sub Text5_CLick()
Text5.Text = ""
End Sub
Private Sub Text6_CLick()
Text6.Text = ""
End Sub
Private Sub Text7_CLick()
Text7.Text = ""
End Sub
Private Sub Text8_CLick()
Text8.Text = ""
End Sub
Private Sub Text9_CLick()
Text9.Text = ""
End Sub
Private Sub Text10_CLick()
Text10.Text = ""
End Sub
Private Sub Text11_CLick()
Text11.Text = ""
End Sub
Private Sub Text12_CLick()
Text12.Text = ""
End Sub
Private Sub Text13_CLick()
Text13.Text = ""
End Sub
Private Sub Text13_Change()
Label9.Caption = Val(Text13.Text)
End Sub
Private Sub Text14_CLick()
Text14.Text = ""
End Sub
Private Sub Text16_CLick()
Text16.Text = ""
End Sub
Private Sub Text17_CLick()
Text17.Text = ""
End Sub
Private Sub Text18_CLick()
Text18.Text = ""
End Sub
Private Sub Text18_Change()
Label10.Caption = Val(Text18.Text)
End Sub
Private Sub Text19_CLick()
Text19.Text = ""
End Sub
Private Sub Text21_CLick()
Text21.Text = ""
End Sub

Private Sub Command1_Click()

a = Val(Text13.Text)
b = Val(Text18.Text)
h = Val(Text1.Text)
Eps = Val(Text2.Text)

N = Abs(b - a) / h + 1
ReDim Arr(N, N)
ReDim Barr(N)
ReDim X(N)
ReDim Y(N)
ReDim P(N + 1)
ReDim Q(N + 1)
ReDim R(N)
ReDim E(N)
ReDim Z(N)

X(0) = a
For i = 1 To N - 1
X(i) = X(i - 1) + h
Next

If Option1.Value = True Then
    For i = 0 To N - 1
        R(i) = Val(Text3.Text) * Val(Text5.Text) * X(i) ^ Val(Text6.Text)
    Next
End If
If Option8.Value = True Then
    For i = 0 To N - 1
        E(i) = Val(Text4.Text) * Val(Text7.Text) * X(i) ^ Val(Text8.Text)
    Next
End If
If Option15.Value = True Then
    For i = 0 To N - 1
        Z(i) = Val(Text9.Text) * Val(Text9.Text) * X(i) ^ Val(Text11.Text)
    Next
End If
If Option2.Value = True Then
    For i = 0 To N - 1
        R(i) = Val(Text3.Text) * Sin(Val(Text5.Text) * X(i) ^ Val(Text6.Text))
    Next
End If
If Option9.Value = True Then
    For i = 0 To N - 1
        E(i) = Val(Text4.Text) * Sin(Val(Text7.Text) * X(i) ^ Val(Text8.Text))
    Next
End If
If Option16.Value = True Then
    For i = 0 To N - 1
        Z(i) = Val(Text9.Text) * Sin(Val(Text9.Text) * X(i) ^ Val(Text11.Text))
    Next
End If
If Option3.Value = True Then
    For i = 0 To N - 1
        R(i) = Val(Text3.Text) * Cos(Val(Text5.Text) * X(i) ^ Val(Text6.Text))
    Next
End If
If Option10.Value = True Then
    For i = 0 To N - 1
        E(i) = Val(Text4.Text) * Cos(Val(Text7.Text) * X(i) ^ Val(Text8.Text))
    Next
End If
If Option17.Value = True Then
    For i = 0 To N - 1
        Z(i) = Val(Text9.Text) * Cos(Val(Text9.Text) * X(i) ^ Val(Text11.Text))
    Next
End If
If Option4.Value = True Then
    For i = 0 To N - 1
        R(i) = Val(Text3.Text) * Tan(Val(Text5.Text) * X(i) ^ Val(Text6.Text))
    Next
End If
If Option11.Value = True Then
    For i = 0 To N - 1
        E(i) = Val(Text4.Text) * Tan(Val(Text7.Text) * X(i) ^ Val(Text8.Text))
    Next
End If
If Option18.Value = True Then
    For i = 0 To N - 1
        Z(i) = Val(Text9.Text) * Tan(Val(Text9.Text) * X(i) ^ Val(Text11.Text))
    Next
End If
If Option5.Value = True Then
    For i = 0 To N - 1
        R(i) = Val(Text3.Text) / Tan(Val(Text5.Text) * X(i) ^ Val(Text6.Text))
    Next
End If
If Option12.Value = True Then
    For i = 0 To N - 1
        E(i) = Val(Text4.Text) / Tan(Val(Text7.Text) * X(i) ^ Val(Text8.Text))
    Next
End If
If Option19.Value = True Then
    For i = 0 To N - 1
        Z(i) = Val(Text9.Text) / Tan(Val(Text9.Text) * X(i) ^ Val(Text11.Text))
    Next
End If
If Option6.Value = True Then
    For i = 0 To N - 1
        R(i) = Val(Text3.Text) * Log(Val(Text5.Text) * X(i) ^ Val(Text6.Text))
    Next
End If
If Option13.Value = True Then
    For i = 0 To N - 1
        E(i) = Val(Text4.Text) * Log(Val(Text7.Text) * X(i) ^ Val(Text8.Text))
    Next
End If
If Option20.Value = True Then
    For i = 0 To N - 1
        Z(i) = Val(Text9.Text) * Log(Val(Text9.Text) * X(i) ^ Val(Text11.Text))
    Next
End If
If Option7.Value = True Then
    For i = 0 To N - 1
        R(i) = Val(Text3.Text) * Exp(Val(Text5.Text) * X(i) ^ Val(Text6.Text))
    Next
End If
If Option14.Value = True Then
    For i = 0 To N - 1
        E(i) = Val(Text4.Text) * Exp(Val(Text7.Text) * X(i) ^ Val(Text8.Text))
    Next
End If
If Option21.Value = True Then
    For i = 0 To N - 1
        Z(i) = Val(Text9.Text) * Exp(Val(Text9.Text) * X(i) ^ Val(Text11.Text))
    Next
End If

Alpha = Val(Text12.Text)
Beta = Val(Text14.Text)
Fi = Val(Text16.Text)
Hamma = Val(Text17.Text)
Delta = Val(Text19.Text)
Psi = Val(Text21.Text)

For i = 0 To N - 1
    For j = 0 To N - 1
        Arr(i, j) = 0
    Next
    Barr(i) = 0
Next

Arr(0, 0) = 1
Arr(0, 1) = Alpha / (h * Beta - Alpha)
Barr(0) = Fi * h / (h * Beta - Alpha)
Arr(N - 1, N - 2) = -Hamma / (h * Delta + Hamma)
Arr(N - 1, N - 1) = 1
Barr(N - 1) = Psi * h / (h * Delta + Hamma)

For i = 1 To N - 2
Arr(i, i - 1) = 2 - h * R(i)
Arr(i, i) = 2 * (h ^ 2) * E(i) - 4
Arr(i, i + 1) = 2 + h * R(i)
Barr(i) = 2 * (h ^ 2) * Z(i)
Next

P(0) = 0
Q(0) = 0
P(1) = -Arr(0, 1) / Arr(0, 0)
Q(1) = Barr(0) / Arr(0, 0)

For i = 2 To N - 1
P(i) = -Arr(i - 1, i) / (Arr(i - 1, i - 1) + Arr(i - 1, i - 2) * P(i - 1))
Q(i) = (Barr(i - 1) - Arr(i - 1, i - 2) * Q(i - 1)) / (Arr(i - 1, i - 1) + Arr(i - 1, i - 2) * P(i - 1))
Next
P(N) = 0
Q(N) = (Barr(N - 1) - Arr(N - 1, N - 2) * Q(N - 1)) / (Arr(N - 1, N - 1) + Arr(N - 1, N - 2) * P(N - 2))

For i = N - 1 To 0 Step -1
Y(i) = P(i + 1) * Y(i + 1) + Q(i + 1)
Next

For i = 0 To N - 1
str = str & X(i) & vbTab & Round(Y(i), Eps) & vbCrLf
Next

Dim Alpha1(2), Alpha2(3), Ar(3, 3), br(3), max As Double
N = N - 1

For i = 0 To 2
    For j = 0 To 2
        Ar(i, j) = 0
    Next
    br(i) = 0
Next

Ar(0, 0) = N + 1
For i = 0 To N
    Ar(0, 1) = Ar(0, 1) + X(i)
    Ar(1, 1) = Ar(1, 1) + X(i) ^ 2
    Ar(1, 2) = Ar(1, 2) + X(i) ^ 3
    Ar(2, 2) = Ar(2, 2) + X(i) ^ 4
    br(0) = br(0) + Y(i)
    br(1) = br(1) + Y(i) * X(i)
    br(2) = br(2) + Y(i) * X(i) ^ 2
Next
Ar(0, 2) = Ar(1, 1)
Ar(1, 0) = Ar(0, 1)
Ar(2, 0) = Ar(1, 1)
Ar(2, 1) = Ar(1, 2)

'Нахождение квадратичного многочлена:
For k = 0 To 2
max = Abs(Ar(k, k))
  iTemp = k
   For i = k + 1 To 2
     If Abs(Ar(i, k)) > max Then
       max = Abs(Ar(i, k))
       iTemp = i
     End If
   Next
   
   For j = 0 To 2
    Temp = Ar(k, j)
    Ar(k, j) = Ar(iTemp, j)
    Ar(iTemp, j) = Temp
   Next
    Temp = br(k)
    br(k) = br(iTemp)
    br(iTemp) = Temp

If Ar(k, k) <> 0 Then
    Temp = Ar(k, k)
    For j = 0 To 2
    Ar(k, j) = Ar(k, j) / Temp
    Next
    br(k) = br(k) / Temp
End If

    For i = k + 1 To 2
        If Ar(k, k) <> 0 Then
        Temp = Ar(i, k) / Ar(k, k)
            For j = k To 2
                Ar(i, j) = Ar(i, j) - Ar(k, j) * Temp
            Next
            br(i) = br(i) - br(k) * Temp
        End If
    Next
Next
   
str = str & vbCrLf & "Квадратичный многочлен по МНК:" & vbCrLf
Alpha2(2) = br(2) / Ar(2, 2)
    For i = 1 To 0 Step -1
        Temp = 0
        For j = i + 1 To 2
            Temp = Temp + Ar(i, j) * Alpha2(j)
        Next
        Alpha2(i) = (br(i) - Temp) / Ar(i, i)
    Next
str = str & Round(Alpha2(0), Eps + 1) & "+" & Round(Alpha2(1), Eps + 1) & "*x+" & Round(Alpha2(2), Eps + 1) & "*x^2" & vbCrLf
Text15.Text = str

Picture1.ScaleMode = vbPixels
Picture1.BackColor = RGB(255, 255, 255)
  
dx = Abs((X(N) - X(0))) / Picture1.ScaleWidth
  
Dim min As Double
max = Y(0)
min = Y(N)
    For i = X(0) To X(N) - dx Step dx
         If max < Alpha2(0) + Alpha2(1) * i + Alpha2(2) * i ^ 2 Then
            max = Alpha2(0) + Alpha2(1) * i + Alpha2(2) * i ^ 2
         End If
         If min > Alpha2(0) + Alpha2(1) * i + Alpha2(2) * i ^ 2 Then
            min = Alpha2(0) + Alpha2(1) * i + Alpha2(2) * i ^ 2
         End If
     Next
  
If X(0) < X(N) Then
    Picture1.Scale (X(0), max)-(X(N), min)
Else
    Picture1.Scale (X(N), max)-(X(0), min)
End If

  
Picture1.Line (X(0), 0)-(X(N), 0)
Picture1.Line (0, max)-(0, min)
  For i = X(0) To X(N) - dx Step dx
      Dim func1, func2 As Double
      func1 = Alpha2(0) + Alpha2(1) * i + Alpha2(2) * i ^ 2
      func2 = Alpha2(0) + Alpha2(1) * (i + dx) + Alpha2(2) * (i + dx) ^ 2
      Picture1.Line (i, func1)-(i + dx, func2), RGB(0, 0, 0)
  Next
  
sFile = ".\Output.txt"
    Open sFile For Output As #1
    Print #1, Text15.Text
    Close #1

End Sub
