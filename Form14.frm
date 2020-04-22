VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H00FF00FF&
   Caption         =   "Cuboid"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14475
   LinkTopic       =   "Form14"
   ScaleHeight     =   9405
   ScaleWidth      =   14475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   4200
      Width           =   9015
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   3360
      Width           =   9015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   5
      Top             =   6960
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3600
      TabIndex        =   4
      Top             =   5640
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   2520
      Width           =   9015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF00FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   1440
      Width           =   9015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Go To Main Menu"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   1
      Top             =   5640
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF00FF&
      Caption         =   "Volume of Cuboid-----------"
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
      Index           =   1
      Left            =   720
      TabIndex        =   11
      Top             =   4320
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF00FF&
      Caption         =   "Enter Height of Cuboid--------"
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
      Left            =   480
      TabIndex        =   10
      Top             =   3360
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF00FF&
      Caption         =   "Enter Bredth of Cuboid-------"
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
      Left            =   480
      TabIndex        =   9
      Top             =   2520
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF00FF&
      Caption         =   "Enter Length of Cuboid-------"
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
      Index           =   0
      Left            =   480
      TabIndex        =   8
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Volume of Cuboid"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   -120
      Width           =   14175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Volume of Cuboid"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14295
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.SetFocus
Dim a As Double
Dim b As Double
a = Text1.Text
b = Text2.Text
c = Text3.Text
Text4.Text = a * b * c
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
calculator.Show
End Sub
