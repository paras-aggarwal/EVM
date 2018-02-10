VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   13800
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   3435
      TabIndex        =   16
      Top             =   3240
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   13800
      Picture         =   "Form2.frx":2B88
      ScaleHeight     =   2955
      ScaleWidth      =   3435
      TabIndex        =   15
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9720
      TabIndex        =   14
      Top             =   6840
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1095
      Left            =   5400
      Top             =   6840
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1931
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\evm\voter.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\evm\voter.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "search"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   13
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   5160
      TabIndex        =   7
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label12 
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   5520
      Width           =   3615
   End
   Begin VB.Label Label11 
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   4680
      Width           =   3615
   End
   Begin VB.Label Label10 
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   3840
      Width           =   3615
   End
   Begin VB.Label Label9 
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label8 
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label Label7 
      Caption         =   "d.o.b"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "address"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "gender"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "father name"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "name"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000B&
      Caption         =   "vote no"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "                        Search vote"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim f As Integer
Private Sub Command1_Click()
Dim gen As String

If Len(Text1.Text) = 0 Then
MsgBox ("enter vote no")
Text1.SetFocus
Exit Sub
End If
Adodc1.Recordset.MoveFirst
f = 0
On Error GoTo xx
While Adodc1.Recordset.EOF = False
If Adodc1.Recordset.Fields("voteno").Value = Text1.Text Then
Label8.Caption = Adodc1.Recordset.Fields("name").Value
Label9.Caption = Adodc1.Recordset.Fields("fname").Value
Label10.Caption = Adodc1.Recordset.Fields("gender").Value
gen = Adodc1.Recordset.Fields("gender").Value

Label11.Caption = Adodc1.Recordset.Fields("city").Value
Label12.Caption = Adodc1.Recordset.Fields("dob").Value
Image1.Picture = LoadPicture(Text1.Text & ".jpg")
Image1.Visible = True

Exit Sub
f = 1
End If
Adodc1.Recordset.MoveNext
Wend
If f = 0 Then
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = ""
Image1.Visible = False

MsgBox ("your vote no is wrong")
End If
Exit Sub
xx:

If gen = "MALE" Then
Image1.Picture = LoadPicture("male.jpg")
Image1.Visible = True
Else
Image1.Picture = LoadPicture("female.jpg")
Image1.Visible = True
End If
MsgBox ("Exact Snap not Available")




End Sub

Private Sub Command2_Click()
Unload Me
Form5.Show
End Sub

