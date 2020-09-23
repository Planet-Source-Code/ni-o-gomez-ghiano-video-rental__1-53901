VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form logincust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "logincust.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5250
   Begin MSAdodcLib.Adodc ado 
      Height          =   375
      Left            =   840
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\nunobone\videorental\video.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\nunobone\videorental\video.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from customerinfo"
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
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Loggin"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "Customer Username"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      Caption         =   "Customer Password / ID#"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   120
      Picture         =   "logincust.frx":0442
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "My Only Bidyo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   """Your No. 1 Stop For Your Video Needs"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "logincust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim datenow As String
Set rs = New ADODB.Recordset
rs.Open "customerinfo", conn, adOpenKeyset, adLockPessimistic, adCmdTable
If Not rs.BOF Then
rs.MoveFirst
End If
Do Until rs.EOF
If rs.Fields("fname") = Text1.Text And rs.Fields("id") = Text2.Text Then
custlog = True
frmmainmdi.Caption = "M.O.B. Video" + "   Welcome  " + rs("fname") + " " + rs("lname")
clname = rs("lname")
cfname = Text1.Text
cidno = ""
cidno = Text2.Text
'cfname = rs("fname")
'cidno = rs("id")
rs.Fields("enter") = "yes"
rs.Fields("date") = Now()
If dueid = rs("id") Then
rs.Fields("due") = duecost
End If
rs.Update
Unload Me
Exit Sub
'Else
'MsgBox "Invalid Username or Password", , "Customer Loggin"
Exit Sub
End If

rs.MoveNext
Loop


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Customer Loggin"

Text1.Text = ""
Text2.Text = ""
End Sub
