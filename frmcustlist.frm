VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcustlist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer List - MOB v1.0"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11430
   Icon            =   "frmcustlist.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   11430
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   8280
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cmbserchby 
      Height          =   315
      ItemData        =   "frmcustlist.frx":0442
      Left            =   3720
      List            =   "frmcustlist.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtserch 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton cmdvi 
      Caption         =   "View Info"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   9600
      TabIndex        =   7
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton command3 
      Caption         =   "Search"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   4575
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   15
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "id"
         Text            =   "ID#"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "lname"
         Text            =   "Last Name"
         Object.Width           =   3052
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "fname"
         Text            =   "First Name"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "mname"
         Text            =   "M.I."
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "address"
         Text            =   "Address"
         Object.Width           =   6085
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "city"
         Text            =   "City"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Postal"
         Text            =   "Postal"
         Object.Width           =   1446
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Due"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Search By:"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmcustlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iclick As Boolean
Private Sub cmbserchby_Click()
If cmbserchby.ListIndex = 0 Then
txtserch.Text = "Enter ID Number"

txtserch.SetFocus
ElseIf cmbserchby.ListIndex = 1 Then
txtserch.Text = "Enter Last Name"
ElseIf cmbserchby.ListIndex = 2 Then
txtserch.Text = "Enter First Name"
End If
End Sub

Private Sub cmddelete_Click()
On Error GoTo diri
If iclick = True Then
For x = 1 To lv1.ListItems.Count
    If lv1.ListItems(x).Selected = True Then
        If MsgBox("Are you sure you want to delete?", vbYesNo, "DELETE CUSTOMER") = vbYes Then
            'id # sa selected nga item
            indicate = lv1.ListItems(x).Text
            lv1.ListItems.Remove (x)
            
        End If
    End If
Next
If Not rs.BOF Then
rs.MoveFirst
End If
Do Until rs.EOF
    If rs.Fields("id") = indicate Then
        rs.Delete
        rs.MoveNext
        Exit Sub
    End If
rs.MoveNext
Loop
End If

diri:

End Sub

Private Sub cmdedit_Click()
If custlog = False Then
For x = 1 To lv1.ListItems.Count
    If lv1.ListItems(x).Selected = True Then
    indicate = lv1.ListItems(x).Text
    editcust.Show
    End If
Next
Else
MsgBox "Please Exit Customer Before Editing", , "My Only Bidyo"
End If
End Sub

Private Sub cmdexit_Click()
frmmainmdi.Toolbar1.Buttons(3).Enabled = False
Unload Me
End Sub

Private Sub Command3_Click()
If cmbserchby.ListIndex = -1 Then

ElseIf cmbserchby.ListIndex = 0 Then
For x = 1 To lv1.ListItems.Count
    If txtserch.Text = lv1.ListItems(x).Text Then
    MsgBox "Found!!!" + vbNewLine + "Customer Id#: " + lv1.ListItems(x).Text, , "Video FOUND!"
    lv1.ListItems(x).Selected = True
    lv1.SetFocus
    End If
Next
Else
For x = 1 To lv1.ListItems.Count
    If txtserch.Text = lv1.ListItems(x).ListSubItems(cmbserchby.ListIndex) Then 'i.ListSubItems(1) Then
        MsgBox "Found!!!" + vbNewLine + "Customer Id#: " + lv1.ListItems(x).Text, , "Video FOUND!"
        lv1.ListItems(x).Selected = True
        lv1.SetFocus
    End If
Next
End If

End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
rs.Open "customerinfo", conn, adOpenKeyset, adLockPessimistic, adCmdTable
cmddelete.Enabled = False
If Not rs.BOF Then
rs.MoveFirst
End If
Do Until rs.EOF
Set lv = lv1.ListItems.Add(, , rs("id"))
lv.SubItems(1) = rs("lname")
lv.SubItems(2) = rs("fname")
lv.SubItems(3) = rs("mname")
lv.SubItems(4) = rs("address")
lv.SubItems(5) = rs("city")
lv.SubItems(6) = rs("postal")
rs.MoveNext
Loop
For x = 1 To lv1.ListItems.Count
    lv1.ListItems(x).Selected = False
Next

'sets the button edit customer in the toolbar as true
'kung ang frmcustlist nga form mo "load"
frmmainmdi.Toolbar1.Buttons(3).Enabled = True
End Sub



Private Sub Form_Unload(Cancel As Integer)
frmmainmdi.Toolbar1.Buttons(3).Enabled = False
End Sub

Private Sub lv1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lv1.SortKey = ColumnHeader.Index - 1
lv1.Sorted = True

End Sub


Private Sub lv1_ItemClick(ByVal Item As MSComctlLib.ListItem)
iclick = True
cmddelete.Enabled = True

End Sub
