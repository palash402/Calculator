VERSION 5.00
Begin VB.Form Form18 
   BackColor       =   &H0000FFFF&
   Caption         =   "Cone"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14580
   LinkTopic       =   "Form18"
   ScaleHeight     =   9660
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      BackColor       =   &H0000FFFF&
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
      Top             =   4080
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
      Left            =   5520
      TabIndex        =   4
      Top             =   7800
      Width           =   3015
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
      Left            =   2520
      TabIndex        =   3
      Top             =   6360
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H0000FFFF&
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
      TabIndex        =   2
      Top             =   3240
      Width           =   9015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FFFF&
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
      TabIndex        =   1
      Top             =   2160
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
      Top             =   6360
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Caption         =   "Volume of Cone----------------"
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
      Top             =   4080
      Width           =   4695
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Enter Height ------------------"
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
      TabIndex        =   8
      Top             =   3240
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Enter the Radius -------------"
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
      Left            =   600
      TabIndex        =   7
      Top             =   2160
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Volume of Cone"
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
      TabIndex        =   6
      Top             =   0
      Width           =   14535
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim r As Double
Dim h As Double
r = Text1.Text
h = Text2.Text
Text3.Text = (3.1415 * (r * r) * h) * (1 / 3)
Text1.SetFocus
End Sub
