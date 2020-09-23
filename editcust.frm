VERSION 5.00
Begin VB.Form editcust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Customer"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "editcust.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   7470
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4200
      TabIndex        =   13
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   5280
      Width           =   1815
   End
   Begin VB.ComboBox cmbcity 
      Height          =   315
      ItemData        =   "editcust.frx":0442
      Left            =   3960
      List            =   "editcust.frx":0458
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox tpostal 
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
      Left            =   5880
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox tad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "editcust.frx":048B
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox tciti 
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
      Left            =   3720
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox ttel 
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
      Left            =   480
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox tage 
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
      Left            =   5160
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox tyr 
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
      Left            =   3240
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox cmbbday 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Text            =   "cmbbday"
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox cmbmonth 
      Height          =   315
      ItemData        =   "editcust.frx":0491
      Left            =   480
      List            =   "editcust.frx":04B6
      TabIndex        =   3
      Text            =   "cmbmonth"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox tid 
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "txtid"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox tf 
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
      Left            =   3720
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox tm 
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
      Left            =   6720
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox tl 
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
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Address"
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Postal Code"
      Height          =   255
      Left            =   5880
      TabIndex        =   26
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "City"
      Height          =   255
      Left            =   3960
      TabIndex        =   25
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Citizenship"
      Height          =   255
      Left            =   3720
      TabIndex        =   24
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Telephone Number"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Age"
      Height          =   255
      Left            =   5160
      TabIndex        =   22
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Year"
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Day"
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Month"
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "M.I."
      Height          =   255
      Left            =   6720
      TabIndex        =   18
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "First Name"
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Last Name"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "ID #"
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "editcust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmcustlist.cmddelete.Enabled = False
If Not rs.BOF Then
rs.MoveFirst
End If
Do Until rs.EOF
    If rs.Fields("id") = indicate Then
        rs.Fields("fname") = tf.Text
        rs("lname") = tl.Text
        rs("mname") = tm.Text
        rs("month") = cmbmonth.Text
        rs("day") = cmbbday.Text
        rs("year") = tyr.Text
        rs("tel") = ttel.Text
        rs("citi") = tciti.Text
        rs("address") = tad.Text
        rs("city") = cmbcity.Text
        rs("postal") = tpostal.Text
        rs.Update
        End If
rs.MoveNext
Loop
frmcustlist.lv1.ListItems.Clear
If Not rs.BOF Then
rs.MoveFirst
End If
Do Until rs.EOF
Set lv = frmcustlist.lv1.ListItems.Add(, , rs("id"))
lv.SubItems(1) = rs("lname")
lv.SubItems(2) = rs("fname")
lv.SubItems(3) = rs("mname")
lv.SubItems(4) = rs("address")
lv.SubItems(5) = rs("city")
lv.SubItems(6) = rs("postal")
rs.MoveNext
Loop
Unload Me
For x = 1 To frmcustlist.lv1.ListItems.Count
If indicate = frmcustlist.lv1.ListItems(x).Text Then
    frmcustlist.lv1.ListItems(x).Selected = True
    'frmcustlist.lv1.SetFocus
End If
Next
End Sub

Private Sub Command2_Click()
frmcustlist.cmddelete.Enabled = False
For x = 1 To frmcustlist.lv1.ListItems.Count
frmcustlist.lv1.ListItems(x).Selected = False
Next
Unload Me
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
rs.Open "customerinfo", conn, adOpenKeyset, adLockPessimistic, adCmdTable
If Not rs.BOF Then
rs.MoveFirst
End If
Do Until rs.EOF
    If rs.Fields("id") = indicate Then
    tid.Text = rs.Fields("id")
    tl.Text = rs.Fields("lname")
    tf.Text = rs.Fields("fname")
    tm.Text = rs.Fields("mname")
    cmbmonth.Text = rs.Fields("month")
    cmbbday.Text = rs.Fields("day")
    tyr.Text = rs.Fields("year")
    tage.Text = rs.Fields("age")
    ttel.Text = rs.Fields("tel")
    tciti.Text = rs.Fields("citi")
    tad.Text = rs.Fields("address")
    cmbcity.Text = rs.Fields("city")
    tpostal.Text = rs.Fields("postal")
    Exit Sub
    End If
    rs.MoveNext
Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
For x = 1 To frmcustlist.lv1.ListItems.Count
frmcustlist.lv1.ListItems(x).Selected = False
Next
frmcustlist.cmddelete.Enabled = False
End Sub
