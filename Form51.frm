VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF00FF&
   Caption         =   "Prime Factors"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Left            =   360
      TabIndex        =   7
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
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
      Left            =   7440
      TabIndex        =   6
      Top             =   5040
      Width           =   2655
   End
   Begin VB.TextBox txtFactors 
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
      Height          =   615
      Left            =   3000
      TabIndex        =   3
      Top             =   2640
      Width           =   6975
   End
   Begin VB.CommandButton cmdFactor 
      Caption         =   "Factor"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Bodoni"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   2
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtNumber 
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
      Height          =   525
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   6975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Prime Factors"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   9615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF00FF&
      Caption         =   "Prime Factors:--"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      Caption         =   "Number:---------"
      BeginProperty Font 
         Name            =   "Bodoni MT"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   480
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   2790
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Factor the number.
Private Sub cmdFactor_Click()
Dim num As Long
Dim factors As New Collection
Dim factor As Variant
Dim result As Long

    ' Get the number's factors.
    num = Val(txtNumber.Text)
    Set factors = FindFactors(num)

    ' Validate the result (for bug-proofing).
    result = 1
    For Each factor In factors
        result = result * factor
    Next factor
    Debug.Assert (result = num)

    ' Display the factors.
    txtFactors.Text = FactorsToString(factors)
End Sub

' Return the number's prime factors.
Private Function FindFactors(ByVal num As Long) As Collection
Dim result As New Collection
Dim factor As Long
    
    ' Take out the 2s.
    Do While (num Mod 2 = 0)
        result.Add 2
        num = num \ 2
    Loop

    ' Take out other primes.
    factor = 3
    Do While (factor * factor <= num)
        If (num Mod factor = 0) Then
            ' This is a factor.
            result.Add factor
            num = num \ factor
        Else
            ' Go to the next odd number.
            factor = factor + 2
        End If
    Loop

    ' If num is not 1, then whatever is left is prime.
    If (num > 1) Then result.Add num

    Set FindFactors = result
End Function

' Return a string like "2 x 3 x 5".
Private Function FactorsToString(ByVal factors As Collection) As String
Dim result As String
Dim factor As Variant

    result = ""
    For Each factor In factors
        result = result & " x " & Format$(factor)
    Next factor

    ' Remove the initial " x ".
    FactorsToString = Mid$(result, 3)
    txtNumber.SetFocus
    
End Function

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
calculator.Show
End Sub
