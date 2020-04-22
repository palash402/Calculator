VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00FFFF00&
   Caption         =   "Volume Menu"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14550
   LinkTopic       =   "Form12"
   ScaleHeight     =   9180
   ScaleWidth      =   14550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6000
      TabIndex        =   8
      Top             =   6240
      Width           =   3015
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Frustum"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      TabIndex        =   7
      Top             =   4440
      Width           =   3495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Sphere"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10680
      TabIndex        =   6
      Top             =   3000
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pyramid"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      TabIndex        =   5
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cone"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   4
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cylender"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10680
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cuboid"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cube"
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Volume of Shapes"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form13.Show
End Sub

Private Sub Command2_Click()
Form14.Show
End Sub

Private Sub Command3_Click()
Form15.Show
End Sub

Private Sub Command4_Click()
Form21.Show
End Sub

Private Sub Command5_Click()
Form22.Show
End Sub

Private Sub Command6_Click()
Form23.Show
End Sub

Private Sub Command7_Click()
Form24.Show
End Sub

Private Sub Command8_Click()
Unload Me
End Sub
