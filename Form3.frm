VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BackColor       =   &H80000007&
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "back"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   16
      Top             =   6120
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   9480
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\evm\CANDI.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\evm\CANDI.mdb;Persist Security Info=False"
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
   Begin VB.CommandButton Command2 
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   15
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   14
      Top             =   6120
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form3.frx":0000
      Left            =   3480
      List            =   "Form3.frx":0019
      TabIndex        =   12
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Ambala"
      Top             =   3120
      Width           =   2415
   End
   Begin VB.OptionButton Option3 
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "MP"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "MLA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Image Image7 
      Height          =   960
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1560
   End
   Begin VB.Label Label7 
      Caption         =   "Election symbol"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Party"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   10
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Place"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Election type"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000B&
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Name of Candidate"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "                CANDIDATE   LIST"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If Combo1.Text = "CONGRESS" Then
Image7.Picture = LoadPicture("CONGRESS.jpg")
End If

End Sub

Private Sub Combo1_Click()
If Combo1.Text = "CONGRESS" Then
Image7.Picture = LoadPicture("CONGRESS.jpg")
End If
If Combo1.Text = "BJP" Then
Image7.Picture = LoadPicture("BJP.jpg")
End If
If Combo1.Text = "BSP" Then
Image7.Picture = LoadPicture("BSP.jpg")
End If
If Combo1.Text = "INLD" Then
Image7.Picture = LoadPicture("INLD.jpg")
End If
If Combo1.Text = "JC" Then
Image7.Picture = LoadPicture("JC.jpg")
End If
If Combo1.Text = "AAP" Then
Image7.Picture = LoadPicture("AAP.jpg")
End If
If Combo1.Text = "OTHERS" Then
Image7.Picture = LoadPicture("OTHERS.jpg")
End If
End Sub

Private Sub Command1_Click()

If (Len(Text1.Text) = 0) Then
MsgBox ("pls enter the candidate name")
Text1.SetFocus
Exit Sub
End If
If (Len(Text2.Text) = 0) Then
MsgBox ("pls enter the address")
Text2.SetFocus
Exit Sub
End If
If (Len(Text3.Text) = 0) Then
MsgBox ("pls enter the City")
Text3.SetFocus
Exit Sub
End If
If ((Option1.Value = False) And (Option2.Value = False) And (Option3.Value = False)) Then
MsgBox ("pls choose election type")
Exit Sub
End If
Adodc1.Recordset.AddNew
If Option1.Value = True Then
electiontype = "MLA"
End If
If Option2.Value = True Then
electiontype = "MP"
End If
If Option3.Value = True Then
electiontype = "MC"
End If
Adodc1.Recordset.Fields("NameofCandidate").Value = UCase(Text1.Text)
Adodc1.Recordset.Fields("address").Value = UCase(Text2.Text)
Adodc1.Recordset.Fields("place").Value = UCase(Text3.Text)
Adodc1.Recordset.Fields("electiontype").Value = electiontype
Adodc1.Recordset.Fields("party").Value = Combo1.Text

'Adodc1.Recordset.Fields("electionsymbol").Value = electionsymbol

Adodc1.Recordset.Update
MsgBox ("your data is save")
Command1.Enabled = False
Exit Sub
xx:
Adodc1.Recordset.CancelUpdate
Exit Sub





End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
End Sub

Private Sub Command3_Click()
Unload Me
Form5.Show
End Sub
