VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H80000007&
   Caption         =   "Form5"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label7 
      Caption         =   "                           EXIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   6
      Top             =   9600
      Width           =   4335
   End
   Begin VB.Label Label6 
      Caption         =   "                  BALLET PAPER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   5
      Top             =   8400
      Width           =   4335
   End
   Begin VB.Label Label5 
      Caption         =   "             CANDIDATE DETAILS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   7200
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "                SEARCH VOTER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   6000
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "               VOTER DETAILS "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   4920
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "              EVM MAIN MENU"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   3720
      Width           =   6135
   End
   Begin VB.Image Image2 
      Height          =   1680
      Left            =   11400
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2280
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   4440
      Picture         =   "Form5.frx":17C2
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "               election commission of india"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()
Unload Me
Form1.Show
End Sub

Private Sub Label4_Click()
Unload Me
Form2.Show
End Sub

Private Sub Label5_Click()
Unload Me
Form3.Show
End Sub

Private Sub Label6_Click()
Unload Me
Form4.Show
End Sub

Private Sub Label7_Click()
End
End Sub
