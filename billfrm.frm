VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form billfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BILL"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "billfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6150
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   4200
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pay"
      Height          =   615
      Left            =   2520
      TabIndex        =   10
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtmnem 
      BorderStyle     =   0  'None
      DataField       =   "mname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      DataField       =   "fname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtlnem 
      BorderStyle     =   0  'None
      DataField       =   "lname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtidno 
      BorderStyle     =   0  'None
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox dbill 
      DataField       =   "due"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox tbill 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox cbill 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1680
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   4080
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\unzipped\videoproject\video.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\unzipped\videoproject\video.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from customerinfo"
      Caption         =   "ado"
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
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5880
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5880
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label3 
      Caption         =   "Due:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Current bill:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Your Total Bill: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   2295
   End
End
Attribute VB_Name = "billfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click() 'pislit sa pay nga button
If (MsgBox("Sure you want to pay?", vbYesNo, "Pay Bill")) = vbYes Then
Adodc1.Recordset.Fields("due") = 0
Adodc1.Recordset.Update
totcost = 0
tbill.Text = Val(dbill.Text) + Val(cbill.Text)
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=" & App.Path & "\video.mdb"
cbill.Text = totcost
If Not Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveFirst
End If

Do Until Adodc1.Recordset.Fields("id") = cidno 'Adodc1.Recordset.EOF
'If Adodc1.Recordset.Fields("id") = cidno Then
'End If
Adodc1.Recordset.MoveNext
Loop
tbill.Text = Val(dbill.Text) + Val(cbill.Text)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
cbill.Text = totcost
tbill.Text = Val(dbill.Text) + Val(cbill.Text)
End Sub
