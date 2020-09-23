VERSION 5.00
Begin VB.Form adddvd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "newvideo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtactor 
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txtdirector 
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox stocknum 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   615
      Left            =   1080
      TabIndex        =   4
      Top             =   4440
      Width           =   1455
   End
   Begin VB.ComboBox cmbgenre 
      Height          =   315
      ItemData        =   "newvideo.frx":0442
      Left            =   1920
      List            =   "newvideo.frx":0467
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
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
      Left            =   480
      MaxLength       =   100
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Actor"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "Director"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000001&
      Caption         =   "Stock #"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lbltit 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "DVD Add"
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
      Left            =   480
      TabIndex        =   8
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Titlle of Video"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Genre of Video"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      BackStyle       =   1  'Opaque
      Height          =   5175
      Left            =   120
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label G 
      Alignment       =   2  'Center
      Caption         =   "Genre of Video"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Titlle of Video"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "adddvd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Not rs.EOF Then
    rs.MoveLast
End If
If Len(tittle.Text) = 0 Or Len(txtdirector.Text) = 0 Or Len(txtactor.Text) = 0 Or Len(cmbgenre.Text) = 0 Then
MsgBox "Please Enter the required fields!"
Else
rs.AddNew
rs.Fields("stock#") = stocknum.Text
rs.Fields("tittle") = tittle.Text
rs.Fields("genre") = cmbgenre.Text
rs.Fields("dvdstat") = "IN"
rs.Fields("rentedby") = " "
rs.Fields("return") = " "
rs("director") = txtdirector.Text
rs("actor") = txtactor.Text
rs("price") = 15
rs.Update
Unload Me
End If
If idvd = True Then
listdvd.l1.ListItems.Clear
    If Not rs.BOF Then
        rs.MoveFirst
    End If
    With rs
    Do Until .EOF
        Set lv = listdvd.l1.ListItems.Add(, , .Fields("stock#"))
        lv.SubItems(1) = .Fields("tittle")
        lv.SubItems(2) = .Fields("genre")
        lv.SubItems(3) = .Fields("dvdstat")
        lv.SubItems(4) = .Fields("rentedby")
        lv.SubItems(5) = .Fields("return")
        .MoveNext
    Loop
    End With
End If



End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
'kung mo butangan nimo og delete na fuction
'buhata pareha sa newcustomer nimo!!!.. kasabot???
Set rs = New ADODB.Recordset
rs.Open "dvd", conn, adOpenKeyset, adLockPessimistic, adCmdTable
Me.Caption = "Add New DVD"
lbltit.Caption = "ADD  NEW  DVD"

stockcount = rs.RecordCount + 1
stocknum.Text = Format(stockcount, "0000")

End Sub

