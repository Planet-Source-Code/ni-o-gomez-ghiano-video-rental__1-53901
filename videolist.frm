VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form listdvd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Video List"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   Icon            =   "videolist.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   9780
   Begin VB.CommandButton cmdpreview 
      Caption         =   "Print"
      Height          =   495
      Left            =   8160
      TabIndex        =   6
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   8160
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdreturn 
      Caption         =   "Return"
      Height          =   495
      Left            =   8160
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   8160
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdrent 
      Caption         =   "Rent"
      Height          =   495
      Left            =   8160
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2400
      Top             =   4200
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
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
   Begin MSComctlLib.ListView l1 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
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
      NumItems        =   6
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
         Alignment       =   2
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
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "lbl"
      BeginProperty Font 
         Name            =   "Pristina"
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
      Width           =   7575
   End
End
Attribute VB_Name = "listdvd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdpreview_Click()
'printpre.Show
Dim pagecount As Integer
pagecount = 1
If MsgBox("Are You Sure You Want to Print DVD list?", vbYesNo, "Print?") = vbYes Then
If Not rs.BOF Then
rs.MoveFirst
End If
Printer.CurrentX = 10
Printer.CurrentY = 10
Printer.Print "Page: " & pagecount
Printer.CurrentY = 200
Printer.CurrentX = 4500
Printer.FontSize = 13
Printer.FontUnderline = True
Printer.Print "DVD LIST"
Printer.FontUnderline = False


Printer.FontSize = 10
Printer.CurrentX = 1350
Printer.CurrentY = 800
Printer.FontSize = 10

Printer.FontBold = True
Printer.CurrentX = 1350
Printer.Print "Stock#";
Printer.CurrentX = 3000
Printer.Print "Tittle";
Printer.CurrentX = 8500
Printer.Print "Genre"  'notice na wala nay semicolon..
'para increment ang Y axis
Printer.CurrentY = 1300
Printer.FontBold = False
Printer.CurrentX = 1350
Do Until rs.EOF
If Printer.CurrentY >= (Printer.ScaleHeight) - 1000 Then
    Printer.NewPage
    pagecount = pagecount + 1
    Printer.CurrentX = 10
    Printer.CurrentY = 10
    Printer.Print "Page: " & pagecount
       
    Printer.CurrentY = 500
    Printer.FontSize = 10

Printer.FontBold = True
Printer.CurrentX = 1350
Printer.Print "Stock#";
Printer.CurrentX = 3000
Printer.Print "Tittle";
Printer.CurrentX = 8500
Printer.Print "Genre"
Printer.CurrentY = 1000
Printer.FontBold = False
Printer.CurrentX = 1350

End If
Printer.CurrentX = 1350
 Printer.Print rs("stock#");
 Printer.CurrentX = 3000
Printer.Print rs("tittle");
Printer.CurrentX = 8500
Printer.Print rs("genre") + vbNewLine 'notice nasad???
    rs.MoveNext
Loop
Printer.EndDoc
Else
Exit Sub
End If

End Sub

Private Sub cmdsearch_Click()
Dim sercH As String
sercH = InputBox("Enter Tittle to Search: ", "Search DVD")
For x = 1 To l1.ListItems.Count
    If sercH = l1.ListItems(x).ListSubItems(1) Then
        MsgBox "Found!!!" + vbNewLine + "At Stock#: " + l1.ListItems(x).Text, , "Video FOUND!"
        l1.SetFocus
        l1.ListItems(x).Selected = True
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
        rs.MoveNext
    Loop
End Sub

Private Sub l1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'sort
l1.SortKey = ColumnHeader.Index - 1
l1.Sorted = True

End Sub

