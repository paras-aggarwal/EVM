VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form6 
   BackColor       =   &H80000007&
   Caption         =   "Form6"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form6"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      Height          =   735
      Left            =   9600
      TabIndex        =   12
      Top             =   5520
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   690
      Left            =   9480
      Top             =   4320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1217
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
   Begin VB.Label Label12 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   11
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label11 
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label10 
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label9 
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label8 
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label7 
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   735
      Left            =   3120
      Picture         =   "Form6.frx":0000
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Image Image5 
      Height          =   720
      Left            =   3120
      Picture         =   "Form6.frx":28BE
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   1200
   End
   Begin VB.Image Image4 
      Height          =   735
      Left            =   3120
      Picture         =   "Form6.frx":40AC
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   795
      Left            =   3120
      Picture         =   "Form6.frx":5E1E
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1140
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   3120
      Picture         =   "Form6.frx":83D9
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   3120
      Picture         =   "Form6.frx":AF40
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1050
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(6) As Integer
Dim i As Integer
Dim j As Integer
Dim n As String
Dim h As Integer



Private Sub Command1_Click()
Unload Me
Form5.Show
End Sub

Private Sub Form_Activate()

a(1) = Val(Label7.Caption)
a(2) = Val(Label8.Caption)
a(3) = Val(Label9.Caption)
a(4) = Val(Label10.Caption)
a(5) = Val(Label11.Caption)
a(6) = Val(Label12.Caption)
h = 0

If a(1) = a(2) Or a(2) = a(3) Or a(3) = a(4) Or a(5) = a(6) Or a(6) = a(1) Then
MsgBox ("election tie")
Exit Sub
End If


For i = 1 To 6
If a(i) > h Then
h = a(i)
j = i
End If
Next i
MsgBox ("Highest vote : " & h)
If j = 1 Then
MsgBox ("winner :" & Label1.Caption)
n = Label1.Caption
End If
If j = 2 Then
MsgBox ("winner :" & Label2.Caption)
n = Label2.Caption
End If
If j = 3 Then
MsgBox ("winner :" & Label3.Caption)
n = Label3.Caption
End If

If j = 4 Then
MsgBox ("winner :" & Label4.Caption)
n = Label4.Caption
End If

If j = 5 Then
MsgBox ("winner :" & Label5.Caption)
n = Label5.Caption
End If

If j = 6 Then
MsgBox ("winner :" & Label6.Caption)
n = Label6.Caption
End If




End Sub

Private Sub Form_Load()
Label1.Caption = Adodc1.Recordset.Fields("nameofcandidate").Value
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
Label2.Caption = Adodc1.Recordset.Fields("nameofcandidate").Value
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
'If Adodc1.Recordset.Fields("party").Value = "OTHERS" Then
'Image2.Picture = LoadPicture("OTHERS.jpg")
'End If
Adodc1.Recordset.MoveNext
Label3.Caption = Adodc1.Recordset.Fields("nameofcandidate").Value
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
'If Adodc1.Recordset.Fields("party").Value = "OTHERS" Then
'Image3.Picture = LoadPicture("OTHERS.jpg")
'End If
Adodc1.Recordset.MoveNext
Label4.Caption = Adodc1.Recordset.Fields("nameofcandidate").Value
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
'If Adodc1.Recordset.Fields("party").Value = "OTHERS" Then
'Image4.Picture = LoadPicture("OTHERS.jpg")
'End If
Adodc1.Recordset.MoveNext
Label5.Caption = Adodc1.Recordset.Fields("nameofcandidate").Value
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
Label6.Caption = Adodc1.Recordset.Fields("nameofcandidate").Value
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





End Sub

