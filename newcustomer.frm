VERSION 5.00
Begin VB.Form newcustomer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register New Customer"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   Icon            =   "newcustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5865
   ScaleWidth      =   8625
   Begin VB.TextBox cmbage 
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
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox memid 
      Height          =   375
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7080
      TabIndex        =   14
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   5160
      TabIndex        =   13
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txtzipcode 
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
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   12
      Top             =   3960
      Width           =   1335
   End
   Begin VB.ComboBox cmbcity 
      Height          =   315
      ItemData        =   "newcustomer.frx":0442
      Left            =   5880
      List            =   "newcustomer.frx":0458
      TabIndex        =   11
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox txtadd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3840
      Width           =   4455
   End
   Begin VB.TextBox txtciti 
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
      Left            =   3600
      MaxLength       =   25
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txttel 
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
      Left            =   960
      MaxLength       =   30
      TabIndex        =   8
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtbirthyr 
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
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox cmbbday 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox cmbmonth 
      Height          =   315
      ItemData        =   "newcustomer.frx":048B
      Left            =   960
      List            =   "newcustomer.frx":04B0
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtname 
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
      Index           =   2
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtname 
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
      Index           =   1
      Left            =   960
      MaxLength       =   30
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox txtname 
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
      Index           =   0
      Left            =   4320
      MaxLength       =   25
      TabIndex        =   2
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label M 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Member ID:"
      Height          =   375
      Left            =   1200
      TabIndex        =   27
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Zip / Postal Code"
      Height          =   255
      Left            =   7080
      TabIndex        =   25
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   255
      Left            =   5880
      TabIndex        =   24
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   3
      Left            =   960
      Top             =   3720
      Width           =   7695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Citizenship"
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone No."
      Height          =   255
      Left            =   960
      TabIndex        =   21
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   960
      Top             =   2760
      Width           =   7695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      Height          =   255
      Left            =   5880
      TabIndex        =   20
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth  MM - DD - YYYY"
      Height          =   255
      Left            =   960
      TabIndex        =   19
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   960
      Top             =   1800
      Width           =   7695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1095
      Index           =   2
      Left            =   0
      TabIndex        =   18
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1095
      Index           =   1
      Left            =   0
      TabIndex        =   17
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   1095
      Index           =   0
      Left            =   0
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "M.I."
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Last name"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "First name"
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   960
      Top             =   840
      Width           =   7695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000006&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   5895
      Left            =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000003&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   5895
      Left            =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "newcustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbbday_Change()
Call changeage
End Sub

Private Sub cmbbday_Click()
Call changeage
End Sub

Private Sub cmbmonth_Change()
Call changeage
End Sub

Private Sub cmbmonth_Click()
If cmbmonth.ListIndex = 0 Then
cmbbday.Clear
For x = 1 To 31
cmbbday.AddItem x
Next
ElseIf cmbmonth.ListIndex = 1 Then
cmbbday.Clear
For x = 1 To 29
cmbbday.AddItem x
Next
ElseIf cmbmonth.ListIndex = 2 Then
cmbbday.Clear
For x = 1 To 31
cmbbday.AddItem x
Next
ElseIf cmbmonth.ListIndex = 3 Then
cmbbday.Clear
For x = 1 To 30
cmbbday.AddItem x
Next
ElseIf cmbmonth.ListIndex = 4 Then
cmbbday.Clear
For x = 1 To 31
cmbbday.AddItem x
Next
ElseIf cmbmonth.ListIndex = 5 Then
cmbbday.Clear
For x = 1 To 30
cmbbday.AddItem x
Next
ElseIf cmbmonth.ListIndex = 6 Then
cmbbday.Clear
For x = 1 To 31
cmbbday.AddItem x
Next
ElseIf cmbmonth.ListIndex = 7 Then
cmbbday.Clear
For x = 1 To 31
cmbbday.AddItem x
Next
ElseIf cmbmonth.ListIndex = 8 Then
cmbbday.Clear
For x = 1 To 30
cmbbday.AddItem x
Next
ElseIf cmbmonth.ListIndex = 9 Then
cmbbday.Clear
For x = 1 To 31
cmbbday.AddItem x
Next
ElseIf cmbmonth.ListIndex = 10 Then
cmbbday.Clear
For x = 1 To 30
cmbbday.AddItem x
Next
ElseIf cmbmonth.ListIndex = 11 Then
cmbbday.Clear
For x = 1 To 31
cmbbday.AddItem x
Next
End If
Call changeage
End Sub

Private Sub Command1_Click()
If txtbirthyr.Text <> "" Then
    If txtname(0).Text = "" Or txtname(1).Text = "" Or txtname(2).Text = "" Or cmbmonth.Text = "" Or cmbbday.Text = "" Or txtbirthyr.Text = "" Or _
        cmbage.Text = "" Or txttel.Text = "" Or txtciti.Text = "" Or txtadd.Text = "" Or _
        cmbcity.Text = "" Or txtzipcode.Text = "" Then
        MsgBox "Please Enter Required Fields"
    ElseIf Not (IsNumeric(txtbirthyr.Text)) Then
        MsgBox "Year must be Numeric", , "My Only Bidyo"
    ElseIf Not (IsNumeric(txttel.Text)) Then
        MsgBox "Telephone must be Numeric", , "My Only Bidyo"
    ElseIf cmbmonth.ListIndex = 1 And (Val(txtbirthyr.Text) Mod 4) <> 0 And cmbbday.Text = 29 Then
        MsgBox "Invalid Birthday or Birth year. Theres no Feb. 29 in " + txtbirthyr.Text, vbCritical, "Invalid Info!"
    Else
        rs.AddNew
        rs.Fields("id") = memid.Text
        rs.Fields("fname") = txtname(0).Text
        rs.Fields("lname") = txtname(1).Text
        rs.Fields("mname") = txtname(2).Text
        rs.Fields("month") = cmbmonth.Text
        rs.Fields("day") = Val(cmbbday.Text)
        rs.Fields("year") = Val(txtbirthyr.Text)
        rs.Fields("age") = cmbage.Text
        rs.Fields("tel") = txttel.Text
        rs.Fields("citi") = txtciti.Text
        rs.Fields("address") = txtadd.Text
        rs.Fields("city") = cmbcity.Text
        rs.Fields("postal") = txtzipcode.Text
        rs.Update
        Unload Me
    End If
End If

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo here
Dim memnum As Integer
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=" & App.Path & "\video.mdb"
    conn.Open
Set rs = New ADODB.Recordset
rs.Open "customerinfo", conn, adOpenKeyset, adLockPessimistic, adCmdTable
If Not rs.EOF Then
rs.MoveLast
End If
memnum = Val(rs("id")) + 1
memid.Text = Format(memnum, "0000")
here:
memid.Text = Format(memnum, "0000")
End Sub
Private Sub changeage()
Dim y, M, d, tage As Long


If cmbmonth.ListIndex + 1 <= Val(Month(Now())) Then
    If Val(cmbbday.ListIndex + 1) <= Val(Day(Now())) Then
        cmbage.Text = Val(Year(Now())) - Val(txtbirthyr.Text)
    Else
        cmbage.Text = (Val(Year(Now())) - Val(txtbirthyr.Text)) - 1
    End If
ElseIf Val(cmbmonth.ListIndex + 1) >= Val(Month(Now())) Then
cmbage.Text = Year(Now()) - Val(txtbirthyr.Text) - 1
End If

End Sub
Private Sub txtbirthyr_Change()
Call changeage
End Sub

Private Sub txtbirthyr_Click()
Call changeage
End Sub
