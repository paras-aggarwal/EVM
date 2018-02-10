VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H80000007&
   Caption         =   "Form4"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   14160
      Top             =   3240
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   975
      Left            =   11040
      Top             =   9240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1720
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\evm\ado3.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\evm\ado3.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   11040
      Top             =   7800
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\evm\voter.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\evm\voter.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
      Caption         =   "Adodc2"
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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   13080
      TabIndex        =   19
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "search vote"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      TabIndex        =   18
      Top             =   6600
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
      Caption         =   "cast vote"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   17
      Top             =   6480
      Width           =   2895
   End
   Begin VB.CommandButton Command9 
      Caption         =   "back"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   16
      Top             =   9480
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "refresh"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   9480
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "finish"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   9480
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Height          =   615
      Left            =   5880
      Picture         =   "Form4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Height          =   615
      Left            =   5880
      Picture         =   "Form4.frx":05B1
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Height          =   615
      Left            =   5880
      Picture         =   "Form4.frx":0B62
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Height          =   615
      Left            =   5760
      Picture         =   "Form4.frx":1113
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Height          =   615
      Left            =   5760
      Picture         =   "Form4.frx":16C4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   5760
      Picture         =   "Form4.frx":1C75
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1095
      Left            =   10680
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.Image Image7 
      Height          =   2055
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label8 
      Height          =   375
      Left            =   10920
      TabIndex        =   20
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   735
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Image Image5 
      Height          =   855
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Image Image4 
      Height          =   735
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   975
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "     vote no"
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10680
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      DataSource      =   "Adodc3"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "                   BALLOT     PAPER"
      DataSource      =   "Adodc1"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim ctr As Integer
Dim tbjp As Integer
Dim tcongress As Integer
Dim tbsp As Integer
Dim taap As Integer
Dim tinld As Integer
Dim tjc As Integer

Private Sub Command1_Click()
Command1.Picture = LoadPicture("cast.jpg")
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
tbjp = tbjp + 1
End Sub

Private Sub Command10_Click()
'Unload Me
'Form6.Show
'rem pedning for user verificiaation
Adodc3.Recordset.MoveFirst
While Adodc3.Recordset.EOF() = False
If Adodc3.Recordset.Fields("vote").Value = Text1.Text Then
MsgBox ("you have already cast your vote")
Exit Sub
End If
Adodc3.Recordset.MoveNext
Wend
Adodc3.Recordset.AddNew
s = Now()
Adodc3.Recordset.Fields("vote").Value = Text1.Text
Adodc3.Recordset.Fields("date").Value = s
Adodc3.Recordset.Update
Command10.Enabled = False


Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True

End Sub

Private Sub Command11_Click()
If Len(Text1.Text) = 0 Then
MsgBox ("enter the vote no.")
Text1.SetFocus
Exit Sub
End If
Adodc2.Recordset.MoveFirst
f = 0
On Error GoTo err
While Adodc2.Recordset.EOF = False
If Adodc2.Recordset.Fields("voteno").Value = Text1.Text Then
Label8.Caption = Adodc2.Recordset.Fields("name").Value
'gen = Adodc2.Recordset.Fields("gender").Value
'Label11.Caption = Adodc2.Recordset.Fields("city").Value
'Label12.Caption = Adodc2.Recordset.Fields("dob").Value

Image7.Picture = LoadPicture(Text1.Text & ".jpg")
Image7.Visible = True
Command7.Visible = True
Command10.Visible = True
Command1.Value = False

Exit Sub
f = 1
End If
Adodc2.Recordset.MoveNext
Wend
If f = 0 Then
MsgBox ("wrong voteno")
Exit Sub
End If
err:
MsgBox ("Photo not avialable and u cannot cast your vote")
End Sub

Private Sub Command2_Click()
Command2.Picture = LoadPicture("cast.jpg")
Command1.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
tcongress = tcongress + 1
End Sub

Private Sub Command3_Click()
Command3.Picture = LoadPicture("cast.jpg")
Command1.Enabled = False
Command2.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
tbsp = tbsp + 1

End Sub

Private Sub Command4_Click()
Command4.Picture = LoadPicture("cast.jpg")
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
taap = taap + 1

End Sub

Private Sub Command5_Click()
Command5.Picture = LoadPicture("cast.jpg")
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command6.Enabled = False
tinld = tinld + 1
End Sub

Private Sub Command6_Click()
Command6.Picture = LoadPicture("cast.jpg")
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
tjc = tjc + 1
End Sub

Private Sub Command7_Click()
'000Form6.Show
Text1.Text = ""
Label8.Caption = ""
Image7.Visible = False
Command1.Enabled = True
Command1.Picture = LoadPicture("Cir.jpg")
Command2.Enabled = True
Command2.Picture = LoadPicture("Cir.jpg")
Command2.Enabled = True
Command3.Picture = LoadPicture("Cir.jpg")
Command3.Enabled = True
Command3.Picture = LoadPicture("Cir.jpg")
Command4.Enabled = True
Command4.Picture = LoadPicture("Cir.jpg")
Command5.Enabled = True
Command5.Picture = LoadPicture("Cir.jpg")
Command6.Enabled = True
Command6.Picture = LoadPicture("Cir.jpg")
Command10.Visible = False
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command10.Enabled = True
End Sub

Private Sub Command8_Click()
Form6.Label7.Caption = tbjp
Form6.Label8.Caption = tcongress
Form6.Label9.Caption = tbsp
Form6.Label10.Caption = taap
Form6.Label11.Caption = tinld
Form6.Label12.Caption = tjc
Unload Me

Form6.Show

End Sub

Private Sub Command9_Click()
Unload Me
Form5.Show
End Sub

Private Sub Form_Load()
Command8.Enabled = False

x = InputBox("Voting time in Minutes")
Label2.Caption = Adodc1.Recordset.Fields("nameofcandidate").Value
If Adodc1.Recordset.Fields("party").Value = "CONGRESS" Then
Image1.Picture = LoadPicture("CONGRESS.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "BJP" Then
Image1.Picture = LoadPicture("BJP.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "BSP" Then
Image1.Picture = LoadPicture("BSP.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "INLD" Then
Image1.Picture = LoadPicture("INLD.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "JC" Then
Image1.Picture = LoadPicture("JC.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "AAP" Then
Image1.Picture = LoadPicture("AAP.jpg")
End If

Adodc1.Recordset.MoveNext
Label3.Caption = Adodc1.Recordset.Fields("nameofcandidate").Value
If Adodc1.Recordset.Fields("party").Value = "CONGRESS" Then
Image2.Picture = LoadPicture("CONGRESS.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "BJP" Then
Image2.Picture = LoadPicture("BJP.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "BSP" Then
Image2.Picture = LoadPicture("BSP.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "INLD" Then
Image2.Picture = LoadPicture("INLD.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "JC" Then
Image2.Picture = LoadPicture("JC.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "AAP" Then
Image2.Picture = LoadPicture("AAP.jpg")
End If

Adodc1.Recordset.MoveNext
Label4.Caption = Adodc1.Recordset.Fields("nameofcandidate").Value
If Adodc1.Recordset.Fields("party").Value = "CONGRESS" Then
Image3.Picture = LoadPicture("CONGRESS.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "BJP" Then
Image3.Picture = LoadPicture("BJP.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "BSP" Then
Image3.Picture = LoadPicture("BSP.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "INLD" Then
Image3.Picture = LoadPicture("INLD.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "JC" Then
Image3.Picture = LoadPicture("JC.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "AAP" Then
Image3.Picture = LoadPicture("AAP.jpg")
End If

Adodc1.Recordset.MoveNext
Label5.Caption = Adodc1.Recordset.Fields("nameofcandidate").Value
If Adodc1.Recordset.Fields("party").Value = "CONGRESS" Then
Image4.Picture = LoadPicture("CONGRESS.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "BJP" Then
Image4.Picture = LoadPicture("BJP.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "BSP" Then
Image4.Picture = LoadPicture("BSP.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "INLD" Then
Image4.Picture = LoadPicture("INLD.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "JC" Then
Image4.Picture = LoadPicture("JC.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "AAP" Then
Image4.Picture = LoadPicture("AAP.jpg")
End If

Adodc1.Recordset.MoveNext
Label6.Caption = Adodc1.Recordset.Fields("nameofcandidate").Value
If Adodc1.Recordset.Fields("party").Value = "CONGRESS" Then
Image5.Picture = LoadPicture("CONGRESS.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "BJP" Then
Image5.Picture = LoadPicture("BJP.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "BSP" Then
Image5.Picture = LoadPicture("BSP.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "INLD" Then
Image5.Picture = LoadPicture("INLD.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "JC" Then
Image5.Picture = LoadPicture("JC.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "AAP" Then
Image5.Picture = LoadPicture("AAP.jpg")
End If

Adodc1.Recordset.MoveNext
Label7.Caption = Adodc1.Recordset.Fields("nameofcandidate").Value
If Adodc1.Recordset.Fields("party").Value = "CONGRESS" Then
Image6.Picture = LoadPicture("CONGRESS.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "BJP" Then
Image6.Picture = LoadPicture("BJP.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "BSP" Then
Image6.Picture = LoadPicture("BSP.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "INLD" Then
Image6.Picture = LoadPicture("INLD.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "JC" Then
Image6.Picture = LoadPicture("JC.jpg")
End If
If Adodc1.Recordset.Fields("party").Value = "AAP" Then
Image6.Picture = LoadPicture("AAP.jpg")
End If




'//////////////////////

Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command10.Visible = False
Command9.Visible = False

''''''''''''''''''''''''

End Sub


Private Sub Option2_Click()

End Sub

Private Sub Timer1_Timer()
ctr = ctr + 1
y = x * 60
If ctr > y Then
MsgBox ("time in over")
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command10.Visible = False
Command11.Visible = False
Command9.Visible = False
Label8.Visible = False
Label9.Visible = False
Text1.Visible = False
Command8.Visible = True
Command8.Enabled = True
Timer1.Enabled = False
Image7.Visible = False
End If
End Sub
