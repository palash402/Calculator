VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Area of Parallalogram"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14580
   LinkTopic       =   "Form10"
   ScaleHeight     =   9075
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0C0FF&
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
      Left            =   5040
      TabIndex        =   5
      Top             =   3480
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
      Left            =   6960
      TabIndex        =   4
      Top             =   6840
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
      Left            =   4080
      TabIndex        =   3
      Top             =   5160
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0FF&
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
      Top             =   2640
      Width           =   9135
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0FF&
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
      TabIndex        =   1
      Top             =   1560
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
      Left            =   8280
      TabIndex        =   0
      Top             =   5160
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Area is:-------------------------"
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
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Enter Base--------------------"
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
      TabIndex        =   8
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Enter Perpendicular----------"
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
      TabIndex        =   7
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Area of Parallalogram"
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
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   14415
   End
End
Attribute VB_Name = "Form10"
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
Text3.Text = a * b
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
calculator.Show

End Sub

