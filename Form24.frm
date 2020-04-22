VERSION 5.00
Begin VB.Form Form24 
   BackColor       =   &H00FF00FF&
   Caption         =   "Frustum of Cone"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14625
   LinkTopic       =   "Form24"
   ScaleHeight     =   9675
   ScaleWidth      =   14625
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   6
      Top             =   5760
      Width           =   3975
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
      TabIndex        =   5
      Top             =   1560
      Width           =   9015
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
      TabIndex        =   4
      Top             =   2640
      Width           =   9015
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
      TabIndex        =   3
      Top             =   5760
      Width           =   3255
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
      TabIndex        =   2
      Top             =   7080
      Width           =   2775
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
      TabIndex        =   1
      Top             =   3480
      Width           =   9015
   End
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
      TabIndex        =   0
      Top             =   4440
      Width           =   9015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Volume of Frustum of Cone"
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
      TabIndex        =   11
      Top             =   0
      Width           =   14655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF00FF&
      Caption         =   "Enter Upper Radius-----------"
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
      TabIndex        =   10
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF00FF&
      Caption         =   "Enter Lower Radius----------"
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
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF00FF&
      Caption         =   "Enter Height of Frustum-----"
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
      Top             =   3480
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF00FF&
      Caption         =   "Volume of Frustum---------"
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
      TabIndex        =   7
      Top             =   4440
      Width           =   4455
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim r1 As Long
Dim r2 As Long
Dim h As Long
r1 = Text1.Text
r2 = Text2.Text
h = Text3.Text
Text4.Text = h * (1 / 3) * (22 / 7) * ((r1 * r2) + (r1 * r1) + (r2 * r2))
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
calculator.Show
End Sub
