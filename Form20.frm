VERSION 5.00
Begin VB.Form Form20 
   BackColor       =   &H00FFFF00&
   Caption         =   "Trignometry"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14640
   LinkTopic       =   "Form20"
   ScaleHeight     =   9330
   ScaleWidth      =   14640
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   12
      Top             =   2880
      Width           =   7095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   1800
      Width           =   7095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SinA"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   8
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CosA"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   6240
      TabIndex        =   7
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "TanA"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   9960
      TabIndex        =   6
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CosecA"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   5
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "SecA"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6240
      TabIndex        =   4
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CotA"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      TabIndex        =   3
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   7800
      TabIndex        =   0
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Result---------------------------"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   11
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Enter Value of X----------------"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   1800
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Trignometry"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14655
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.SetFocus
Dim n1 As Double
n1 = Text1.Text
Text2.Text = Sin(n1 * (3.14159265) / 180)
Text1.SetFocus
End Sub

Private Sub Command2_Click(Index As Integer)
Dim Answer As Double
Dim a As Integer
a = Text1.Text
converter = (3.14159265) / 180
Answer = Cos(a * converter)
Text2.Text = Answer
If Text1.Text = 90 Then
Text2.Text = 0
Else: Text2.Text = Cos(a * converter)
End If
End Sub

Private Sub Command3_Click(Index As Integer)
Text1.SetFocus
Dim n1 As Double
n1 = Text1.Text
Text2.Text = Tan(n1 * (3.14159265) / 180)
Text2.SetFocus
If Text1.Text = 90 Then
Text2.Text = "VALUE IS NOT DEFINED"
Else: Text2.Text = Tan(n1 * (3.14159265) / 180)
End If
End Sub

Private Sub Command4_Click()
Text1.SetFocus
Dim n1 As Double
n1 = Text1.Text
Text2.Text = 1 / (Sin(n1 * (3.14159265) / 180))
Text2.SetFocus
End Sub

Private Sub Command5_Click()
Text1.SetFocus
Dim n1 As Double
n1 = Text1.Text
Text2.Text = 1 / (Cos(n1 * (3.14159265) / 180))
Text2.SetFocus
End Sub

Private Sub Command6_Click()
Text1.SetFocus
Dim n1 As Double
n1 = Text1.Text
Text2.Text = 1 / (Tan(n1 * (3.14159265) / 180))
Text2.SetFocus
End Sub

Private Sub Command7_Click(Index As Integer)
calculator.Show
End Sub

Private Sub Command8_Click(Index As Integer)
Unload Me
End Sub

Private Sub Command9_Click()


End Sub
