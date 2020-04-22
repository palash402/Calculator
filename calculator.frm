VERSION 5.00
Begin VB.Form calculator 
   BackColor       =   &H00FF00FF&
   Caption         =   "Main Menu"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14805
   LinkTopic       =   "form30"
   ScaleHeight     =   9060
   ScaleWidth      =   14805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command17 
      Caption         =   "Prime Factors"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      TabIndex        =   17
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CommandButton Command16 
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
      Height          =   1095
      Left            =   5400
      TabIndex        =   16
      Top             =   7200
      Width           =   4575
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Milage Consumpsion"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   15
      Top             =   5880
      Width           =   3015
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Cube root"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   14
      Top             =   5880
      Width           =   3135
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Square root"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      TabIndex        =   13
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Inverse"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11760
      TabIndex        =   12
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   11
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Cube"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   10
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Trignometry"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11760
      TabIndex        =   9
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Square"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11760
      TabIndex        =   8
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Factorial"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      TabIndex        =   7
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   6
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Devide"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   11760
      TabIndex        =   4
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Multiply"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Subtraction"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   2
      Top             =   1800
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Addition"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   14175
   End
End
Attribute VB_Name = "calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form50.Show
End Sub

Private Sub Command10_Click()
Form25.Show
End Sub

Private Sub Command11_Click()
Form26.Show
End Sub

Private Sub Command12_Click()
Form27.Show
End Sub

Private Sub Command13_Click()
Form28.Show
End Sub

Private Sub Command14_Click()
Form29.Show
End Sub

Private Sub Command15_Click()
Form68.Show
End Sub

Private Sub Command16_Click()
Unload Me
End Sub

Private Sub Command17_Click()
Form1.Show
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Form3.Show
End Sub

Private Sub Command4_Click()
Form4.Show
End Sub

Private Sub Command5_Click()
Form5.Show
End Sub

Private Sub Command6_Click()
Form12.Show
End Sub

Private Sub Command7_Click()
Form16.Show
End Sub

Private Sub Command8_Click()
Form17.Show
End Sub

Private Sub Command9_Click()
Form20.Show
End Sub
