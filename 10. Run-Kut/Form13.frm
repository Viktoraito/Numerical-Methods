VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14925
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   14925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Caption         =   "График"
      Height          =   3975
      Left            =   7200
      TabIndex        =   18
      Top             =   240
      Width           =   4575
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   3615
         Left            =   120
         ScaleHeight     =   237
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   285
         TabIndex        =   19
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Решение"
      Height          =   3975
      Left            =   3360
      TabIndex        =   13
      Top             =   240
      Width           =   3495
      Begin VB.TextBox Text1 
         Height          =   3375
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   14
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Начальные настройки"
      Height          =   3975
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      Begin VB.CommandButton Command1 
         Caption         =   "Пуск"
         Height          =   855
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Frame Frame5 
         Caption         =   "Шаг и точность"
         Height          =   735
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1935
         Begin VB.TextBox Text7 
            Height          =   375
            Left            =   1200
            TabIndex        =   8
            Text            =   "0.01"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   360
            TabIndex        =   7
            Text            =   "0.1"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Е="
            Height          =   255
            Left            =   960
            TabIndex        =   12
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "h="
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   255
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Начальные условия"
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   840
            TabIndex        =   6
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Y(0)="
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Интервал решения"
         Height          =   735
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1935
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1200
            TabIndex        =   5
            Text            =   "1"
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Text            =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "b="
            Height          =   255
            Left            =   960
            TabIndex        =   16
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "a="
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   375
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   2400
      Index           =   1
      Left            =   12000
      Picture         =   "Form13.frx":0000
      Top             =   240
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   $"Form13.frx":0BE6
      Height          =   1215
      Index           =   1
      Left            =   12000
      TabIndex        =   17
      Top             =   2880
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A, B, h, Eps, X(), Y(), Z(), delta As Double, str As String, N As Integer

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Text2_Change()
Label1(0).Caption = "Y(" & Val(Text2.Text) & ")="
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

Private Sub Text6_Click()
Text6.Text = ""
End Sub

Private Sub Text7_Click()
Text7.Text = ""
End Sub

Private Sub Command1_Click()
Text1.Text = ""
str = ""
A = Val(Text2.Text)
B = Val(Text3.Text)
h = Val(Text6.Text)
Eps = -Log(Val(Text7.Text)) / Log(10)
N = (B - A) / h
ReDim X(N + 1)
ReDim Y(N + 1)
X(0) = A
For i = 1 To N
X(i) = X(i - 1) + h
Next
Y(0) = Val(Text5.Text)

Dim sFile As String, sWhole As String, v As Variant
    
    sFile = ".\Input.txt"
    Open sFile For Input As #1
    sWhole = Input$(LOF(1), 1)
    Close #1
    v = Split(sWhole, vbNewLine & " ")

For Each Item In v
str = str & Item
Next

str = str & vbCrLf

Dim k1, k2, k3, k4 As Double
For i = 1 To N
k1=((X(i-1))+(Y(i-1)))/((X(i-1))+Exp((Y(i-1))))
k2=((X(i-1)+h/2)+(Y(i-1)+h/2*k1))/((X(i-1)+h/2)+Exp((Y(i-1)+h/2*k1)))
k3=((X(i-1)+h/2)+(Y(i-1)+h/2*k2))/((X(i-1)+h/2)+Exp((Y(i-1)+h/2*k2)))
k4=((X(i-1)+h)+(Y(i-1)+h*k3))/((X(i-1)+h)+Exp((Y(i-1)+h*k3)))
Y(i) = Y(i - 1) + (h / 6) * (k1 + 2 * k2 + 2 * k3 + k4)
Next

delta = Abs(Y(0))
For i = 1 To N
    If delta < Abs(Y(i)) Then delta = Abs(Y(i))
Next

If delta / 15 > Exp(-Eps * Log(10)) Then

Do
h = h / 2
N = (B - A) / h
ReDim X(N + 1)
ReDim Z(N + 1)
X(0) = A
For i = 1 To N
X(i) = X(i - 1) + h
Next
Z(0) = Val(Text5.Text)
For i = 1 To N
k1=((X(i-1))+(Z(i-1)))/((X(i-1))+Exp((Z(i-1))))
k2=((X(i-1)+h/2)+(Z(i-1)+h/2*k1))/((X(i-1)+h/2)+Exp((Z(i-1)+h/2*k1)))
k3=((X(i-1)+h/2)+(Z(i-1)+h/2*k2))/((X(i-1)+h/2)+Exp((Z(i-1)+h/2*k2)))
k4=((X(i-1)+h)+(Z(i-1)+h*k3))/((X(i-1)+h)+Exp((Z(i-1)+h*k3)))
Z(i) = Z(i - 1) + (h / 6) * (k1 + 2 * k2 + 2 * k3 + k4)
Next

delta = Abs(Y(0) - Z(0))
For i = 1 To N / 2
    If delta < Abs(Y(i) - Z(i * 2)) Then delta = Abs(Y(i) - Z(i))
Next

ReDim Y(N + 1)
For i = 0 To N
    Y(i) = Z(i)
Next

Loop While delta / 15 > Exp(-Eps * Log(10))
End If

For i = 0 To N
str = str & "X=" & Text1.Text & X(i) & vbTab & "Y(X)=" & Round(Y(i), Eps + 1) & vbCrLf
Next

Dim Alpha1(2), Alpha2(3), Ar(3, 3), br(3), max As Double

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
str = str & Round(Alpha2(0), Eps + 1) & "+" & Round(Alpha2(1), Eps + 1) & "*x+" & Round(Alpha2(2), Eps + 1) & "*x^2"
Text1.Text = str

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
    Print #1, Text1.Text
    Close #1

End Sub
