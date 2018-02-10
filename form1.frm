VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   22
      Top             =   9840
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   3375
      Left            =   11160
      Picture         =   "form1.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   3435
      TabIndex        =   21
      Top             =   3960
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   11160
      Picture         =   "form1.frx":2B88
      ScaleHeight     =   2955
      ScaleWidth      =   3435
      TabIndex        =   20
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "view"
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
      Left            =   9840
      TabIndex        =   19
      Top             =   9840
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   8280
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
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
      BackColor       =   16744703
      ForeColor       =   -2147483637
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\evm\voter.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\evm\voter.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin MSACAL.Calendar Calendar1 
      Height          =   1815
      Left            =   4200
      TabIndex        =   18
      Top             =   5640
      Width           =   4455
      _Version        =   524288
      _ExtentX        =   7858
      _ExtentY        =   3201
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2016
      Month           =   3
      Day             =   19
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   -2147483639
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "REFRESH"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      TabIndex        =   15
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   14
      Top             =   9840
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FF80FF&
      Caption         =   "FEMALE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   5880
      TabIndex        =   13
      Top             =   7680
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF80FF&
      Caption         =   "MALE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   4200
      TabIndex        =   11
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "AMBALA"
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox TEXT1 
      BackColor       =   &H00FF80FF&
      DataSource      =   "Adodc1"
      ForeColor       =   &H8000000B&
      Height          =   405
      Left            =   4200
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   4320
      Picture         =   "form1.frx":434A
      Stretch         =   -1  'True
      Top             =   8400
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   5880
      Picture         =   "form1.frx":6DB2
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   135
      Left            =   16200
      TabIndex        =   17
      Top             =   5400
      Width           =   135
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF80FF&
      Caption         =   "   VOTER      ELECTION       CARD"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   1080
      TabIndex        =   16
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF80FF&
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF80FF&
      Caption         =   "D.O.B."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF80FF&
      Caption         =   "Father name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF80FF&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF80FF&
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF80FF&
      Caption         =   "Ward no"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Voter no"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo xx
If (Len(Text1.Text) <> 6) Then
MsgBox ("pls enter the voter no")
Text1.SetFocus
Exit Sub
End If
If (Len(Text2.Text) <> 2) Then
MsgBox ("pls enter the ward no")
Text2.SetFocus
Exit Sub
End If
If (Len(Text3.Text) = 0) Then
MsgBox ("pls enter the City")
Text3.SetFocus
Exit Sub
End If
If (Len(Text4.Text) = 0) Then
MsgBox ("pls enter the Name")
Text4.SetFocus
Exit Sub
End If
If (Len(Text5.Text) = 0) Then
MsgBox ("pls enter the Father name")
Text5.SetFocus
Exit Sub
End If
If ((Option1.Value = False) And (Option2.Value = False)) Then
MsgBox ("pls choose gender")
Option1.SetFocus
Exit Sub
End If
Adodc1.Recordset.AddNew
If Option1.Value = True Then
GENDER = "male"
Else
GENDER = "female"
End If
Adodc1.Recordset.Fields("voteno").Value = UCase(Text1.Text)
Adodc1.Recordset.Fields("wardno").Value = UCase(Text2.Text)
Adodc1.Recordset.Fields("city").Value = UCase(Text3.Text)
Adodc1.Recordset.Fields("name").Value = UCase(Text4.Text)
Adodc1.Recordset.Fields("fname").Value = UCase(Text5.Text)
Adodc1.Recordset.Fields("dob").Value = Calendar1.Value

Adodc1.Recordset.Fields("gender").Value = UCase(GENDER)
'dob = calender1.Value
Adodc1.Recordset.Update
MsgBox ("your data is save")
Command1.Enabled = False
Exit Sub
xx:
Adodc1.Recordset.CancelUpdate

MsgBox ("DUPLICATE VOTER NO")
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
'Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Option1.Value = False
Option2.Value = False
Command1.Enabled = True
Image2.Visible = True
Image1.Visible = True


End Sub

Private Sub Command3_Click()
Form2.Show

End Sub

Private Sub Command4_Click()
Unload Me
Form5.Show
End Sub

Private Sub Option1_Click()
Image1.Visible = False
Image2.Visible = True

End Sub

Private Sub Option2_Click()
Image2.Visible = False
Image1.Visible = True

End Sub

