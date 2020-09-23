VERSION 5.00
Begin VB.Form addvcd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "addvcd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   5970
   Begin VB.TextBox txtactor 
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtdirector 
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox tittle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      MaxLength       =   50
      TabIndex        =   4
      Top             =   720
      Width           =   4815
   End
   Begin VB.ComboBox cmbgenre 
      Height          =   315
      ItemData        =   "addvcd.frx":0442
      Left            =   2040
      List            =   "addvcd.frx":0467
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox stocknum 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Actor"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Director"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "ADD NEW VCD"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000001&
      Caption         =   "Stock #"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Titlle of Video"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label G 
      Alignment       =   2  'Center
      Caption         =   "Genre of Video"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   5055
      Left            =   240
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Genre of Video"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Titlle of Video"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label lbltit 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000001&
      Caption         =   "Stock #"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
End
Attribute VB_Name = "addvcd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not rs.EOF Then
    rs.MoveLast
End If
If Len(tittle.Text) = 0 Or Len(cmbgenre.Text) = 0 Then
MsgBox "Please Enter the required fields!"
Else
rs.AddNew
rs.Fields("stock#") = Format(stockcount, "0000")
rs.Fields("tittle") = tittle.Text
rs.Fields("genre") = cmbgenre.Text
rs.Fields("status") = "IN"
rs.Fields("rentedby") = " "
rs.Fields("return") = " "
rs("director") = txtdirector.Text
rs("actor") = txtactor.Text
rs("price") = 15
rs.Update
If ivcd = True Then
listvcd.l1.ListItems.Clear
    If Not rs.BOF Then
        rs.MoveFirst
    End If
    With rs
    Do Until .EOF
        Set lv = listvcd.l1.ListItems.Add(, , .Fields("stock#"))
        lv.SubItems(1) = .Fields("tittle")
        lv.SubItems(2) = .Fields("genre")
        lv.SubItems(3) = .Fields("status")
        lv.SubItems(4) = .Fields("rentedby")
        lv.SubItems(5) = .Fields("return")
        .MoveNext
    Loop
    End With
End If
Unload Me
End If


End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
rs.Open "vcd", conn, adOpenKeyset, adLockPessimistic, adCmdTable
Me.Caption = "Add New VCD"
lbltit.Caption = "ADD  NEW  VCD"
stockcount = rs.RecordCount + 1
stocknum.Text = Format(stockcount, "0000")
Me.Caption = "ADD NEW VCD"
End Sub
