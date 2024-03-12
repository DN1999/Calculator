VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command17 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   19
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   18
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   17
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton Command15 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   16
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command14 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   15
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   14
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   13
      Top             =   5280
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   12
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   11
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   10
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   9
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   6
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   5
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5655
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   4575
      Begin VB.CommandButton Command20 
         Caption         =   "Sqr."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   21
         Top             =   4560
         Width           =   855
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double
Dim b As Double
Dim Operator As Integer
Dim ans As Double
Private Sub Command1_Click()
If Text1.Text <> "0" Then
 Text1.Text = Text1.Text + "1"
Else
 Text1.Text = "1"
End If
End Sub

Private Sub Command12_Click()
a = Val(Text1.Text)
Operator = 3 '* operator
Text1.Text = ""
End Sub

Private Sub Command14_Click()
Dim x As Integer
x = Val(Text1.Text)
If x = Val(Text1.Text) Then
  Text1.Text = Str(x) + "."
End If
End Sub

Private Sub Command15_Click()
If Operator = 1 Then
 b = Val(Text1.Text)
 ans = a + b
 Text1.Text = ans
ElseIf Operator = 2 Then
 b = Val(Text1.Text)
 ans = a - b
 Text1.Text = ans
ElseIf Operator = 3 Then
 b = Val(Text1.Text)
 ans = a * b
 Text1.Text = ans
ElseIf Operator = 4 Then
 b = Val(Text1.Text)
 ans = a / b
 Text1.Text = ans
End If
End Sub

Private Sub Command16_Click()
a = Val(Text1.Text)
Operator = 4 '/ operator
Text1.Text = ""
End Sub

Private Sub Command17_Click()
Text1.Text = ""
End Sub
Private Sub Command18_Click()
End
End Sub

Private Sub Command19_Click()
Dim x As Integer
x = Len(Text1.Text)
If x <> 0 Then
 Text1.Text = Left(Text1.Text, x - 1)
End If
End Sub

Private Sub Command2_Click()
If Text1.Text <> "0" Then
 Text1.Text = Text1.Text + "2"
Else
 Text1.Text = "2"
End If
End Sub

Private Sub Command20_Click()
a = Val(Text1.Text)
ans = a * a
Text1.Text = Str(ans)
End Sub

Private Sub Command3_Click()
If Text1.Text <> "0" Then
 Text1.Text = Text1.Text + "3"
Else
 Text1.Text = "3"
End If
End Sub

Private Sub Command4_Click()
a = Val(Text1.Text)
Operator = 1 '+ operator
Text1.Text = ""
End Sub

Private Sub Command5_Click()
If Text1.Text <> "0" Then
 Text1.Text = Text1.Text + "4"
Else
 Text1.Text = "4"
End If
End Sub
Private Sub Command6_Click()
If Text1.Text <> "0" Then
 Text1.Text = Text1.Text + "5"
Else
 Text1.Text = "5"
End If
End Sub
Private Sub Command7_Click()
If Text1.Text <> "0" Then
 Text1.Text = Text1.Text + "6"
Else
 Text1.Text = "6"
End If
End Sub

Private Sub Command8_Click()
a = Val(Text1.Text)
Operator = 2 '- operator
Text1.Text = ""
End Sub

Private Sub Command9_Click()
If Text1.Text <> "0" Then
 Text1.Text = Text1.Text + "7"
Else
 Text1.Text = "7"
End If
End Sub
Private Sub Command10_Click()
If Text1.Text <> "0" Then
 Text1.Text = Text1.Text + "8"
Else
 Text1.Text = "8"
End If
End Sub
Private Sub Command11_Click()
If Text1.Text <> "0" Then
 Text1.Text = Text1.Text + "9"
Else
 Text1.Text = "9"
End If
End Sub

Private Sub Command13_Click()
If Text1.Text <> "0" Then
 Text1.Text = Text1.Text + "0"
End If
End Sub

Private Sub Form_Load()
Text1.Text = ""
Frame1.Caption = "Calculator"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
KeyAscii = 0
End If
End Sub
