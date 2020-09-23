VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form listdvd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Video List"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   Icon            =   "listdvd.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   11445
   Begin MSAdodcLib.Adodc adodvd 
      Height          =   375
      Left            =   480
      Top             =   4080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.CommandButton cmdview 
      Caption         =   "View Info"
      Height          =   375
      Left            =   9960
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdpreview 
      Caption         =   "Print Preview"
      Height          =   615
      Left            =   9960
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   9960
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return"
      Height          =   495
      Left            =   9960
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   9960
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdrent 
      Caption         =   "Rent"
      Height          =   495
      Left            =   9960
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin MSComctlLib.ListView l1 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7011
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "lvwstockno"
         Text            =   "Stock#"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "lvwtitle"
         Text            =   "Title"
         Object.Width           =   5203
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "lvwgenre"
         Text            =   "Genre"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "lvwstatus"
         Text            =   "Status"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "lvwrented"
         Text            =   "Rented By"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "lvwreturn"
         Text            =   "return"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "lvwactor"
         Text            =   "Actor"
         Object.Width           =   2205
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "lvwdirector"
         Text            =   "Director"
         Object.Width           =   2205
      EndProperty
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "lbl"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "listdvd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdpreview_Click()
DVDreport.Show

'DataEnvironment1.myConnection
'DVDreport.DataSource = "Select * from dvd where tittle like = '%a%'"
End Sub

Private Sub cmdrent_Click()
Set rs = New ADODB.Recordset
rs.Open "dvd", conn, adOpenKeyset, adLockPessimistic, adCmdTable

For X = 1 To l1.ListItems.Count
    If l1.ListItems(X).Selected = True Then
    renti = l1.ListItems(X).Text
    End If
Next
If Not rs.BOF Then
rs.MoveFirst
End If
Do Until rs.EOF
If rs("stock#") = renti Then
    If rs("dvdstat") = "IN" Then
        If MsgBox("Are You sure you want to rent " + "'" + rs("tittle") + "'", vbYesNo, "Rent DVD?") = vbYes Then
        rs("dvdstat") = "OUT"
        totcost = totcost + 15
        rs("rentedby") = cidno
        dvdrents = dvdrents + 1
        rs("return") = DateValue(Now + 3) 'now() + Day(Now() + 3)
        Call here
        Exit Sub
        End If
    Else
    MsgBox "Already Been Rented", vbCritical, "WARNING!"
    End If
    
End If
rs.MoveNext

Loop




End Sub
Private Sub here()

l1.ListItems.Clear
If Not rs.BOF Then
rs.MoveFirst
End If
Do Until rs.EOF
Set lv = l1.ListItems.Add(, , rs("stock#"))
lv.SubItems(1) = rs("tittle")
lv.SubItems(2) = rs("genre")
lv.SubItems(3) = rs("dvdstat")
lv.SubItems(4) = rs("rentedby")
lv.SubItems(5) = rs("return")
lv.SubItems(6) = rs("actor")
lv.SubItems(7) = rs("director")
rs.MoveNext
Loop
'For x = 0 To l1.ListItems.Count
'    If l1.ListItems(x).ListSubItems(4).Text <> "" Then
'    l1.ListItems(x).ListSubItems(5).Text =
'    End If
'l1.ListItems(5).ListSubItems

End Sub

Private Sub cmdreturn_Click()
Set rs = New ADODB.Recordset
rs.Open "dvd", conn, adOpenKeyset, adLockPessimistic, adCmdTable

For X = 1 To l1.ListItems.Count
    If l1.ListItems(X).Selected = True Then
    renti = l1.ListItems(X).Text
    End If
Next
If Not rs.BOF Then
rs.MoveFirst
End If
Do Until rs.EOF
If rs("stock#") = renti Then
    If rs("dvdstat") = "OUT" Then
        If rs("rentedby") = cidno Then
            If MsgBox("Are You sure you want to return " + "'" + rs("tittle") + "'", vbYesNo, "Rent DVD?") = vbYes Then
            rs("dvdstat") = "IN"
            rs("rentedby") = " "
            rs("return") = " "
            Call here
            'bag -o
            If totcost <> 0 Then
            totcost = totcost - 15
            Exit Do
            Exit Sub
            End If
            End If
        Else
            MsgBox "Customer not match!", vbInformation, "MOB"
        End If
    Else
    MsgBox "DVD not RENTED", vbCritical, "WARNING!"
    End If
    
End If
If rs.EOF Then
    Exit Do
    End If
rs.MoveNext
  
Loop

End Sub

Private Sub cmdsearch_Click()
Dim sercH As String
Set rs = New ADODB.Recordset
rs.Open "dvd", conn, adOpenKeyset, adLockPessimistic, adCmdTable

sercH = InputBox("Enter Tittle to Search: ", "Search DVD")
For X = 1 To l1.ListItems.Count
    If sercH = l1.ListItems(X).ListSubItems(1) Then
        MsgBox "Found!!! " + l1.ListItems(X).ListSubItems(1) + vbNewLine + "At Stock#: " + l1.ListItems(X).Text, , "Video FOUND!"
        l1.SetFocus
        l1.ListItems(X).Selected = True
        End If
Next
End Sub

Private Sub cmdview_Click()
For X = 1 To l1.ListItems.Count
    If l1.ListItems(X).Selected = True Then
    MsgBox "Title: " + l1.ListItems(X).ListSubItems(1) + vbNewLine + "Actor: " + l1.ListItems(X).ListSubItems(6) + vbNewLine + "Director: " + l1.ListItems(X).ListSubItems(7), , "View Info"
    End If
Next

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
idvd = True
Set rs = New ADODB.Recordset
If custlog = False Then 'rent command button
    cmdrent.Enabled = False
    cmdreturn.Enabled = False
Else
    cmdreturn.Enabled = True
    cmdrent.Enabled = True
End If

lbl.Caption = "DVD List"
Me.Caption = "DVD List"
rs.Open "dvd", conn, adOpenKeyset, adLockPessimistic, adCmdTable
        If Not rs.BOF Then
        rs.MoveFirst
        End If
    Do Until rs.EOF
        Set lv = l1.ListItems.Add(, , rs.Fields("stock#"))
        lv.SubItems(1) = rs.Fields("tittle")
        lv.SubItems(2) = rs.Fields("genre")
        lv.SubItems(3) = rs.Fields("dvdstat")
        lv.SubItems(4) = rs.Fields("rentedby")
        lv.SubItems(5) = rs.Fields("return")
        lv.SubItems(6) = rs("actor")
        lv.SubItems(7) = rs("director")
        rs.MoveNext
    Loop
End Sub

Private Sub l1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'sort
l1.SortKey = ColumnHeader.Index - 1
l1.Sorted = True

End Sub

Private Sub l1_ItemClick(ByVal Item As MSComctlLib.ListItem)
'asdfasdf

End Sub

Private Sub lbl5_Click()
End Sub
