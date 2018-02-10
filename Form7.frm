VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H80000007&
   Caption         =   "Form7"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   9
      Top             =   8520
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "SURENDER KUMAR "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   8
      Top             =   6480
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "MUNISH GUPTA(H.O.D)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      TabIndex        =   7
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label Label7 
      Caption         =   "SUBMITTED TO"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   6
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "ANSHI YADAV"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "AARTI VERMA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "AARTI DEVI"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "DEEPSHIKHA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "SUBMITTED BY"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   5640
      Picture         =   "Form7.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "  ELECTRONIC VOTING MACHINE  (EVM)"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click1()
Unload Me
Form5.Show
End Sub

Private Sub Command1_Click()
Unload Me
Form5.Show

End Sub
